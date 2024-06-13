class TestCorrectionException(Exception): pass
class DuplicateException(TestCorrectionException): pass
class UnknownElementException(TestCorrectionException): pass
class UndefinedElementException(TestCorrectionException): pass
