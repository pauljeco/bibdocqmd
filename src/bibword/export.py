"""Compatibility import for scripts using `bibword.export`."""

from bibdocqmd.export import *  # noqa: F401,F403


if __name__ == '__main__':
    raise SystemExit(main())
