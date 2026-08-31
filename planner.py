"""
planner.py
----------
Finds a sequence of breeding steps that turns the fish you currently own
(your aquarium) into a desired target fish (body, fin) - or, if only one of
the two attributes is specified, into ANY fish that matches that attribute.

Algorithm:
1. Generational expansion (BFS-like): starting from the fish you own, every
   pair (including a fish bred with itself) is tried and their offspring is
   computed. Every newly discovered fish type keeps track of ALL the ways it
   was produced (parent A, parent B, body rarity, fin rarity, generation).
2. Once the target is reachable (or already owned), we build up to N
   alternative "recipes" (breeding orders), ranked first by rarity (Common
   paths first, then Uncommon, then Rare) and then by length.
3. Each recipe is a list of steps in the order they must be performed
   (parents before offspring), reusing already-owned or already-planned fish
   whenever possible.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Set, Tuple

from excel_loader import MapData, RARITY_ORDER

Fish = Tuple[str, str]  # (body, fin)


@dataclass
class Production:
    parent_a: Fish
    parent_b: Fish
    body_rarity: str
    fin_rarity: str
    generation: int

    @property
    def worst_rarity(self) -> str:
        # "weakest link" - the worse (rarer) of body/fin for this cross
        if RARITY_ORDER[self.body_rarity] >= RARITY_ORDER[self.fin_rarity]:
            return self.body_rarity
        return self.fin_rarity


@dataclass
class BreedStep:
    parent_a: Fish
    parent_b: Fish
    result: Fish
    body_rarity: str
    fin_rarity: str

    @property
    def worst_rarity(self) -> str:
        if RARITY_ORDER[self.body_rarity] >= RARITY_ORDER[self.fin_rarity]:
            return self.body_rarity
        return self.fin_rarity


@dataclass
class Recipe:
    steps: List[BreedStep]  # in execution order
    overall_rarity: str  # worst rarity across the whole recipe ("weakest link")
    result_fish: Fish  # the concrete fish this recipe actually produces

    def describe(self) -> str:
        lines = []
        for i, s in enumerate(self.steps, start=1):
            lines.append(
                f"{i}. {s.parent_a[0]} / {s.parent_a[1]}  x  "
                f"{s.parent_b[0]} / {s.parent_b[1]}  ->  "
                f"{s.result[0]} / {s.result[1]}   "
                f"[body: {s.body_rarity}, fin: {s.fin_rarity}]"
            )
        return "\n".join(lines)


class BreedingGraph:
    def __init__(self, map_data: MapData, owned: List[Fish], max_generations: int = 10,
                 max_pool_size: int = 20000):
        self.map_data = map_data
        self.owned: Set[Fish] = set(owned)
        # productions[fish] = every discovered way to produce that fish
        self.productions: Dict[Fish, List[Production]] = {}
        self.generation_found: Dict[Fish, int] = {f: 0 for f in self.owned}
        self._build(max_generations, max_pool_size)

    def _build(self, max_generations: int, max_pool_size: int):
        pool: Set[Fish] = set(self.owned)

        for gen in range(1, max_generations + 1):
            if len(pool) > max_pool_size:
                break
            pool_list = list(pool)
            newly_found: Set[Fish] = set()

            for i in range(len(pool_list)):
                for j in range(i, len(pool_list)):
                    fa = pool_list[i]
                    fb = pool_list[j]
                    bred = self.map_data.breed_fish(fa, fb)
                    if bred is None:
                        continue
                    result, (body_rarity, fin_rarity) = bred

                    prod = Production(
                        parent_a=fa,
                        parent_b=fb,
                        body_rarity=body_rarity,
                        fin_rarity=fin_rarity,
                        generation=gen,
                    )
                    self.productions.setdefault(result, []).append(prod)

                    if result not in self.generation_found:
                        self.generation_found[result] = gen
                        newly_found.add(result)

            if not newly_found:
                # closure reached - no point simulating further generations
                break
            pool |= newly_found

    def is_reachable(self, fish: Fish) -> bool:
        return fish in self.owned or fish in self.productions

    def _matches(self, fish: Fish, target_body: Optional[str], target_fin: Optional[str]) -> bool:
        if target_body is not None and fish[0] != target_body:
            return False
        if target_fin is not None and fish[1] != target_fin:
            return False
        return True

    def best_recipes(self, target_body: Optional[str] = None, target_fin: Optional[str] = None,
                      top_n: int = 5) -> List[Recipe]:
        """
        Builds up to top_n alternative recipes, ranked by rarity (Common first)
        then by number of steps.

        - If both target_body and target_fin are given: alternative recipes for
          that exact fish (different breeding orders that reach the same fish).
        - If only one is given: alternative recipes across different concrete
          fish that satisfy the given attribute (any fin, or any body).
        """
        if target_body is None and target_fin is None:
            return []

        exact = target_body is not None and target_fin is not None

        if exact:
            target = (target_body, target_fin)
            if target in self.owned:
                return [Recipe(steps=[], overall_rarity="Common", result_fish=target)]
            if target not in self.productions:
                return []

            seen_pairs: Set[frozenset] = set()
            unique_candidates: List[Production] = []
            for p in sorted(
                self.productions[target],
                key=lambda p: (RARITY_ORDER[p.worst_rarity], p.generation),
            ):
                pair_key = frozenset((p.parent_a, p.parent_b))
                if pair_key in seen_pairs:
                    continue
                seen_pairs.add(pair_key)
                unique_candidates.append(p)
            candidates = unique_candidates[:top_n]

            recipes: List[Recipe] = []
            for cand in candidates:
                steps: List[BreedStep] = []
                visited: Set[Fish] = set()
                ok = self._expand(cand.parent_a, steps, visited)
                ok = ok and self._expand(cand.parent_b, steps, visited)
                if not ok:
                    continue
                steps.append(
                    BreedStep(
                        parent_a=cand.parent_a,
                        parent_b=cand.parent_b,
                        result=target,
                        body_rarity=cand.body_rarity,
                        fin_rarity=cand.fin_rarity,
                    )
                )
                recipes.append(
                    Recipe(steps=steps, overall_rarity=self._overall_rarity(steps), result_fish=target)
                )

            recipes.sort(key=lambda r: (RARITY_ORDER[r.overall_rarity], len(r.steps)))
            return recipes

        # Partial target: match across all reachable fish, one recipe per
        # distinct matching fish (using its single best known production).
        scored: List[Tuple[Fish, int, str, bool]] = []  # (fish, generation, rarity, already_owned)

        for fish in self.owned:
            if self._matches(fish, target_body, target_fin):
                scored.append((fish, 0, "Common", True))

        for fish, prods in self.productions.items():
            if fish in self.owned:
                continue
            if self._matches(fish, target_body, target_fin):
                best = sorted(prods, key=lambda p: (RARITY_ORDER[p.worst_rarity], p.generation))[0]
                scored.append((fish, best.generation, best.worst_rarity, False))

        scored.sort(key=lambda t: (0 if t[3] else 1, RARITY_ORDER[t[2]], t[1]))
        scored = scored[:top_n]

        recipes = []
        for fish, _, _, already_owned in scored:
            if already_owned:
                recipes.append(Recipe(steps=[], overall_rarity="Common", result_fish=fish))
                continue
            steps: List[BreedStep] = []
            visited: Set[Fish] = set()
            if self._expand(fish, steps, visited):
                recipes.append(
                    Recipe(steps=steps, overall_rarity=self._overall_rarity(steps), result_fish=fish)
                )

        recipes.sort(key=lambda r: (RARITY_ORDER[r.overall_rarity], len(r.steps)))
        return recipes

    def _overall_rarity(self, steps: List[BreedStep]) -> str:
        worst = "Common"
        for s in steps:
            if RARITY_ORDER[s.worst_rarity] > RARITY_ORDER[worst]:
                worst = s.worst_rarity
        return worst

    def _expand(self, fish: Fish, steps: List[BreedStep], visited: Set[Fish]) -> bool:
        """Recursively adds the steps needed to produce `fish` (if not already
        owned) before the step currently being appended. Uses the best (least
        rare, then shortest) known production for each intermediate fish."""
        if fish in self.owned:
            return True
        if fish in visited:
            # already scheduled earlier in this recipe (shared intermediate)
            return True

        prods = self.productions.get(fish)
        if not prods:
            return False

        best = sorted(prods, key=lambda p: (RARITY_ORDER[p.worst_rarity], p.generation))[0]

        if not self._expand(best.parent_a, steps, visited):
            return False
        if not self._expand(best.parent_b, steps, visited):
            return False

        visited.add(fish)
        steps.append(
            BreedStep(
                parent_a=best.parent_a,
                parent_b=best.parent_b,
                result=fish,
                body_rarity=best.body_rarity,
                fin_rarity=best.fin_rarity,
            )
        )
        return True


def plan_breeding(map_data: MapData, owned: List[Fish], target_body: Optional[str] = None,
                   target_fin: Optional[str] = None, top_n: int = 5,
                   max_generations: int = 10) -> List[Recipe]:
    graph = BreedingGraph(map_data, owned, max_generations=max_generations)
    return graph.best_recipes(target_body=target_body, target_fin=target_fin, top_n=top_n)
