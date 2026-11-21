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


class VBAObjectError(VBARuntimeError, AttributeError):
    """Run-time error 91 or 424: a member used on a value that holds no object (x.Value where x is Nothing or
    Empty). Python meets it as a missing attribute, so it is an AttributeError too."""


class VBACompileError(VBAError):
    """VBA would refuse to compile this code. Never trapped by On Error."""


class VBAParseError(VBAError):
    """The vb2py grammar could not parse this code."""


class VBANotImplemented(VBAError):
    """The code uses something the converter does not model yet.

    Never trapped by On Error: a gap in the converter must stop the run instead of
    being skipped silently.
    """


class VBATimeout(VBAError):
    """The run passed the time it was given. Never trapped by On Error."""


class VBAEnd(Exception):
    """The VBA `End` statement: stop everything."""
