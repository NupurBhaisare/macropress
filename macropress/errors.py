"""Errors raised while parsing or running VBA."""


class VBAError(Exception):
    """Base class for every error the converter raises on purpose."""


class VBARuntimeError(VBAError):
    """A VBA run-time error (Err.Number / Err.Description).

    These are the errors that `On Error Resume Next` and `On Error GoTo` can trap.
    """

    def __init__(self, number, description=""):
        super().__init__(f"run-time error {number}: {description}")
        self.number = number
        self.description = description


class VBACompileError(VBAError):
    """VBA would refuse to compile this code. Never trapped by On Error."""


class VBAParseError(VBAError):
    """The vb2py grammar could not parse this code."""


class VBANotImplemented(VBAError):
    """The code uses something the converter does not model yet.

    Never trapped by On Error: a gap in the converter must stop the run instead of
    being skipped silently.
    """


class VBAEnd(Exception):
    """The VBA `End` statement: stop everything."""
