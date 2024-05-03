class FileNotFoundError(Exception):
    """Raised when the file is not found."""

    def __init__(self, message):
        self.msg = message

    def __str__(self):
        return f"Could not find file: {self.msg}"


class SlideCouldNotBeCopiedError(Exception):
    """Raised when a slide could not be copied."""

    def __init__(self, from_file_path, slide_num, new_file_path, new_slide_num, e):
        self.from_file_path = from_file_path
        self.slide_num = slide_num
        self.new_file_path = new_file_path
        self.new_slide_num = new_slide_num
        self.e = e

    def __str__(self):
        return f'Slide 【{self.slide_num}】 from file "{self.from_file_path}" could not be copied to "{self.new_file_path}" as slide 【{self.new_slide_num}】 already exists. Error: {self.e}'


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


class SlideCouldNotBeReadError(Exception):
    """Raised when a slide could not be read."""

    def __init__(self, file_path, slide_num, e):
        self.file_path = file_path
        self.slide_num = slide_num
        self.e = e

    def __str__(self):
        return f'Slide 【{self.slide_num}】 from file "{self.file_path}" could not be read. Error: {self.e}'


class SlideCouldNotBeSavedAsOtherFormatError(Exception):
    """Raised when a slide could not be saved as other format."""

    def __init__(self, file_path, slide_num, to_format, e):
        self.file_path = file_path
        self.slide_num = slide_num
        self.to_format = to_format
        self.e = e

    def __str__(self):
        return f'Slide 【{self.slide_num}】 from file "{self.file_path}" could not be saved as {self.to_format}. Error: {self.e}'


class TextCouldNotBeChangedError(Exception):
    """Raised when a text could not be changed."""

    def __init__(self, file_path, slide_num, text, new_text, error):
        self.file_path = file_path
        self.slide_num = slide_num
        self.text = text
        self.new_text = new_text
        self.error = error

    def __str__(self):
        return f'Text "{self.text}" in slide 【{self.slide_num}】 from file "{self.file_path}" could not be changed to "{self.new_text}". Error: {self.error}'


class PictureCouldNotBeAddedError(Exception):
    """Raised when a picture could not be added."""

    def __init__(self, file_path, slide_num, image_path, left, top, error):
        self.file_path = file_path
        self.slide_num = slide_num
        self.image_path = image_path
        self.left = left
        self.top = top
        self.error = error

    def __str__(self):
        return f'Picture "{self.image_path}" in slide 【{self.slide_num}】 from file "{self.file_path}" could not be added. Error: {self.error}'


class PictureCouldNotBeChangedError(Exception):
    """Raised when a picture could not be changed."""

    def __init__(self, file_path, slide_num, old_image_bytes, new_image_bytes, error):
        self.file_path = file_path
        self.slide_num = slide_num
        self.old_image_bytes = old_image_bytes
        self.new_image_bytes = new_image_bytes
        self.error = error

    def __str__(self):
        return f'Picture in slide 【{self.slide_num}】 from file "{self.file_path}" could not be changed. Error: {self.error}'
