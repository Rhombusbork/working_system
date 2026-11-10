"""송년회 기숙사 배정: 그룹(교수님·선배 기수·20기)마다 네 기숙사에 같은 인원을 배정한다.

사용법
    python3 sorting_hat.py responses.csv [-o result.csv] [--seed 2026]
    python3 sorting_hat.py --demo                  # 가짜 응답으로 시험
    --questions questions_likert.json              # 리커트 척도판(답은 1~5)

응답 CSV 열: 이름,구분,q1,q2,q3,q4,q5  (답은 A~D, 리커트판은 1~5)
"""
import argparse
import csv
import json
import random
from pathlib import Path

HOUSES = ["G", "H", "R", "S"]


def load(name):
    global QUESTIONS, NAMES
    QUESTIONS = json.loads((Path(__file__).parent / name).read_text(encoding="utf-8"))
    NAMES = QUESTIONS["houses"]


load("questions.json")


def score(answers):
    """답 5개 → 기숙사별 점수."""
    if QUESTIONS.get("type") == "likert":
        return score_likert(answers)
    total = dict.fromkeys(HOUSES, 0)
    for q, a in zip(QUESTIONS["questions"], answers):
        for h, s in q["options"][a]["score"].items():
            total[h] += s
    return total


def score_likert(answers):
    """축마다 (답-3)/2에 역채점 부호를 곱해 평균 → -1~1 좌표. 기숙사 점수 = 좌표·모서리 내적 × 10(정수)."""
    axis = {"x": [], "y": []}
    for q, a in zip(QUESTIONS["questions"], answers):
        axis[q["axis"]].append(q["key"] * (int(a) - 3) / 2)
    x, y = (sum(v) / len(v) for v in (axis["x"], axis["y"]))
    return {h: round(10 * (x * c["x"] + y * c["y"])) for h, c in QUESTIONS["corners"].items()}


def rank(scores, house):
    """배정된 기숙사가 그 사람에게 몇 순위인가(동점은 같은 순위)."""
    return 1 + sum(scores[h] > scores[house] for h in HOUSES)


def cost(scores, house, noise):
    """그 사람의 최고 점수와의 차이를 제곱한다. 제곱이라 한 사람을 멀리 보내기보다 여럿을 조금씩 옮긴다.
    noise(0~999)는 동점일 때만 순서를 가르는 추첨이다(차이 1점 = 1000)."""
    gap = max(scores.values()) - scores[house]
    return gap * gap * 1000 + noise


def quotas(groups, rng):
    """그룹마다 기숙사 정원. 기본은 n//4. 남는 인원은 지금까지 총인원이 가장 적은 기숙사에 준다.
    그래서 그룹 안에서는 최대 1명 차이, 기숙사 총인원도 최대 1명 차이가 된다."""
    running = dict.fromkeys(HOUSES, 0)
    result = {}
    for g in sorted(groups, key=lambda g: -len(groups[g])):
        people = groups[g]
        base, extra = divmod(len(people), 4)
        # 총인원이 같은 기숙사끼리는 이 그룹에서 1순위로 많이 나온 기숙사를 먼저
        demand = {h: sum(rank(p["scores"], h) == 1 for p in people) for h in HOUSES}
        order = sorted(HOUSES, key=lambda h: (running[h], -demand[h], rng.random()))
        q = {h: base + (h in order[:extra]) for h in HOUSES}
        for h in HOUSES:
            running[h] += q[h]
        result[g] = q
    return result


def min_cost_assign(costs, cap):
    """최소비용 흐름(최단 경로를 반복해 늘림). costs[i][h]: 사람 i를 기숙사 h에 넣는 비용, cap[h]: 정원.
    사람 수와 정원 합이 같다고 가정한다. 수십~수백 명이면 즉시 끝난다."""
    n = len(costs)
    src, sink = n + 4, n + 5
    graph = [[] for _ in range(n + 6)]  # 간선: [to, cap, cost, rev]

    def add(u, v, c, w):
        graph[u].append([v, c, w, len(graph[v])])
        graph[v].append([u, 0, -w, len(graph[u]) - 1])

    for i in range(n):
        add(src, i, 1, 0)
        for k, h in enumerate(HOUSES):
            add(i, n + k, 1, costs[i][h])
    for k, h in enumerate(HOUSES):
        add(n + k, sink, cap[h], 0)

    for _ in range(n):  # 한 번에 한 사람씩 흐름을 늘린다(Bellman-Ford, 음수 역간선 허용)
        dist = [float("inf")] * len(graph)
        prev = [None] * len(graph)
        dist[src] = 0
        for _ in range(len(graph)):
            changed = False
            for u in range(len(graph)):
                if dist[u] == float("inf"):
                    continue
                for j, (v, c, w, _) in enumerate(graph[u]):
                    if c > 0 and dist[u] + w < dist[v]:
                        dist[v], prev[v] = dist[u] + w, (u, j)
                        changed = True
            if not changed:
                break
        v = sink
        while v != src:
            u, j = prev[v]
            graph[u][j][1] -= 1
            graph[v][graph[u][j][3]][1] += 1
            v = u

    return [next(HOUSES[e[0] - n] for e in graph[i] if n <= e[0] < n + 4 and e[1] == 0)
            for i in range(n)]


def sort_people(people, seed=2026):
    """people: [{"이름", "구분", "answers"}] → 각자에 "scores", "기숙사", "순위"를 채워 돌려준다."""
    rng = random.Random(seed)
    for p in people:
        p["scores"] = score(p["answers"])
    groups = {}
    for p in people:
        groups.setdefault(p["구분"], []).append(p)
    caps = quotas(groups, rng)
    for g, members in groups.items():
        costs = [{h: cost(p["scores"], h, rng.randrange(1000)) for h in HOUSES} for p in members]
        for p, h in zip(members, min_cost_assign(costs, caps[g])):
            p["기숙사"], p["순위"] = h, rank(p["scores"], h)
    return people


def naive(people):
    """비교용: 점수가 가장 높은 기숙사로 그냥 보낸다(동점은 앞 순서)."""
    return {id(p): max(HOUSES, key=lambda h: p["scores"][h]) for p in people}


def report(people):
    groups = sorted({p["구분"] for p in people}, key=lambda g: [p["구분"] for p in people].index(g))
    plain = naive(people)
    header = "".join(f"{NAMES[h]:>8}" for h in HOUSES)
    for title, house_of in [("그냥 1순위로 배정", lambda p: plain[id(p)]),
                            ("균등 배정", lambda p: p["기숙사"])]:
        print(f"\n[{title}]{'':6}{header}")
        for g in groups + ["합계"]:
            ms = [p for p in people if g in ("합계", p["구분"])]
            print(f"  {g:<12}" + "".join(f"{sum(house_of(p) == h for p in ms):>10}" for h in HOUSES))
    ranks = [p["순위"] for p in people]
    print("\n균등 배정의 순위: " + " · ".join(f"{r}순위 {ranks.count(r)}명" for r in range(1, 5)))


def demo_people(seed=7):
    """가짜 응답. 실제처럼 치우치게(그리핀도르·래번클로 쪽이 많게) 만든다."""
    rng = random.Random(seed)
    sizes = {"교수님": 10, "선배 기수": 15, "20기": 15}
    people = []
    for g, n in sizes.items():
        for i in range(n):
            people.append({"이름": f"{g}{i + 1}", "구분": g,
                           "answers": rng.choices("ABCD", weights=[4, 2, 3, 1], k=5)
                           if QUESTIONS.get("type") != "likert" else
                           [str(min(5, max(1, round(rng.gauss(3.6 if q["key"] > 0 else 2.6, 1)))))
                            for q in QUESTIONS["questions"]]})
    return people


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("responses", nargs="?")
    ap.add_argument("-o", "--out")
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--questions", default="questions.json")
    args = ap.parse_args()
    load(args.questions)

    if args.demo:
        people = demo_people()
    else:
        with open(args.responses, encoding="utf-8-sig") as f:
            people = [{"이름": r["이름"], "구분": r["구분"],
                       "answers": [r[f"q{i}"].strip().upper() for i in range(1, 6)]}
                      for r in csv.DictReader(f)]
    sort_people(people, args.seed)
    report(people)

    if args.out:
        with open(args.out, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(["이름", "구분", "기숙사", "순위"] + [NAMES[h] for h in HOUSES])
            for p in people:
                w.writerow([p["이름"], p["구분"], NAMES[p["기숙사"]], p["순위"]]
                           + [p["scores"][h] for h in HOUSES])


if __name__ == "__main__":
    main()
