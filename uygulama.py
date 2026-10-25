"""Outlook Firma Rehberi v5 compatibility launcher.

The former Windows COM/Tkinter implementation remains available in the Git
history and the v4.6 release tag. New executions use the cross-platform v5
package on both Windows and macOS.
"""

from outlook_rehberi.__main__ import main


if __name__ == "__main__":
    raise SystemExit(main())
