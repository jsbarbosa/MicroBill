class MicrobillError(Exception):
    """ Microbill base error. """

    def __init__(self, message = "Microbill base error"):
        self.message = message

    def __repr__(self):
        return self.message

    def __str__(self):
        return self.message

class IncompatibleError(MicrobillError):
    """ Error de compatibilidad """
    def __init__(self, message = "Error de compatibilidad"):
        super(IncompatibleError, self).__init__(message)
        self.message = message

    def __repr__(self):
        return self.message

    def __str__(self):
        return self.message
