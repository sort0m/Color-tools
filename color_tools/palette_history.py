"""Compact immutable RGBA snapshots with a shared undo payload budget."""

MAX_UNDO_BYTES = 32 * 1024 * 1024


def pack_colors(colors):
    return colors if isinstance(colors, bytes) else bytes(v for c in colors for v in c)


def unpack_colors(snapshot):
    return [list(snapshot[i:i + 4]) for i in range(0, len(snapshot), 4)]


def record_snapshot(stack, name, colors):
    snapshot = pack_colors(colors)
    # Evict even older entries for this palette if this operation cannot be
    # recorded: otherwise Undo would misleadingly skip over the latest change.
    if len(snapshot) > MAX_UNDO_BYTES:
        kept = [entry for entry in stack if entry[0] != name]
        stack.clear()
        stack.extend(kept)
        return
    stack.append((name, snapshot))
    size = sum(len(entry[1]) for entry in stack)
    while size > MAX_UNDO_BYTES:
        size -= len(stack.popleft()[1])
