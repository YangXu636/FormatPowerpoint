import math
import numpy as np


def softmax(logits):
    """Compute softmax values for each sets of scores in logits."""

    e_logits = np.exp(logits - np.max(logits))
    return e_logits / np.sum(e_logits)


def sigmoid(x):
    """Compute sigmoid value for a given input."""

    return 1 / (1 + math.exp(-x))
