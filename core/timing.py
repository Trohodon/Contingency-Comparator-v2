class Timer:
    """Accumulate time against a threshold."""
    def __init__(self, threshold: float):
        self.threshold = float(threshold)
        self.acc = 0.0

    def reset(self):
        self.acc = 0.0

    def set_threshold(self, threshold: float):
        self.threshold = float(threshold)

    def tick(self, dt: float) -> bool:
        self.acc += dt
        if self.acc >= self.threshold:
            self.acc -= self.threshold
            return True
        return False

    def elapsed(self) -> float:
        return self.acc
