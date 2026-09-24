# Day 1 — Kata C: Flatten the flags (stretch, ~10 min)
# Three boolean flags produce eight possible calls, and some are
# duplicated logic in disguise. Find the duplication, then propose
# a shape that doesn't require testing all eight paths.


def report(data, csv=False, summary=False, verbose=False):
    if csv:
        if summary:
            return to_csv(aggregate(data))
        return to_csv(data)
    if summary:
        if verbose:
            return to_text(aggregate(data), detail=True)
        return to_text(aggregate(data))
    if verbose:
        return to_text(data, detail=True)
    return to_text(data)
