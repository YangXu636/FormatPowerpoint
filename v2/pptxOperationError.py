class FileNotFoundError(Exception):
    """Raised when the file is not found."""

    def __init__(self, message):
        self.msg = message

    def __str__(self):
        return f"Could not find file: {self.msg}"


class SlideCouldNotBeCopiedError(Exception):
    """Raised when a slide could not be copied."""

    def __init__(self, from_file_path, slide_num, new_file_path, new_slide_num):
        self.msg = f'Slide 【{slide_num}】 from file "{from_file_path}" could not be copied to "{new_file_path}" as slide 【{new_slide_num}】 already exists.'

    def __str__(self):
        return self.msg


class FileCouldNotBeCreatedError(Exception):
    """Raised when a file could not be created."""

    def __init__(self, file_path):
        self.file_path = file_path

    def __str__(self):
        return f'File "{self.file_path}" could not be created.'


class FileCouldNotBeOpenedError(Exception):
    """Raised when a file could not be opened."""

    def __init__(self, file_path):
        self.file_path = file_path

    def __str__(self):
        return f'File "{self.file_path}" could not be opened.'


class FileCouldNotBeConvertedError(Exception):
    """Raised when a file could not be converted."""

    def __init__(self, file, to_format):
        self.file = file
        self.to_format = to_format

    def __str__(self):
        return f'File "{self.file}" could not be converted to {self.to_format}.'
