"""Opaque UI color pairs meeting WCAG 2.2 enhanced text contrast (7:1)."""

from .color_math import contrast_ratio

TEXT_CONTRAST = 7.0


def contrast_color(preferred, backgrounds, minimum=TEXT_CONTRAST):
    """Move a preferred RGB color toward black/white until every pair passes.

    Backgrounds must be opaque. Raise if no shared foreground can meet the
    target; callers must then change the backgrounds instead of claiming AAA.
    """
    preferred = tuple(int(round(c)) for c in preferred[:3])
    backgrounds = tuple(tuple(bg[:3]) for bg in backgrounds)
    def passes(rgb):
        return all(contrast_ratio(*rgb, *bg) >= minimum for bg in backgrounds)
    if passes(preferred):
        return [*preferred, 255]
    candidates = []
    for end in (0, 255):
        if not passes((end, end, end)):
            continue
        for step in range(1, 256):
            candidate = tuple(round(c + (end - c) * step / 255) for c in preferred)
            if passes(candidate):
                candidates.append((sum((a-b)**2 for a, b in zip(candidate, preferred)), candidate))
                break
    if not candidates:
        raise ValueError('No single text color reaches the requested contrast on these backgrounds')
    return [*min(candidates)[1], 255]
