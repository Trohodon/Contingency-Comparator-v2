import random

PIECES = ["I", "O", "T", "S", "Z", "J", "L"]


class Bag7:
    """7-bag randomizer."""
    def __init__(self, seed=None):
        self.rng = random.Random(seed)
        self.bag = []
        self._refill()

    def _refill(self):
        self.bag = PIECES[:]
        self.rng.shuffle(self.bag)

    def next(self) -> str:
        if not self.bag:
            self._refill()
        return self.bag.pop()
