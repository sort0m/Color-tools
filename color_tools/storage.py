"""Atomic writes shared by settings and palette exports (no UI dependencies)."""

import os
import tempfile


def atomic_write(path, data):
    """Replace path only after all bytes have been written and flushed.

    Each writer owns a temporary file in the destination directory. Callers
    still serialize writes when the order of their snapshots matters.
    """
    path = os.path.abspath(os.fspath(path))
    if isinstance(data, str):
        data = data.encode('utf-8')
    fd, temporary = tempfile.mkstemp(prefix='.colortools-', suffix='.tmp',
                                     dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, 'wb') as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)
