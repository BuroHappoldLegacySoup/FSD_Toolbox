from dataclasses import dataclass

@dataclass
class TableInfo:
    """
    Dataclass to store information about a table to be extracted and converted.
    
    Attributes:
        main_title (str): The main title to search for in the HTML.
        heading_text (str): The text of the heading to search for in the HTML.
        title (str): The title to be used for the table in the Word document.
    """
    main_title: str
    heading_text: str
    title: str


@dataclass
class ImageInfo:
    """Dataclass that stores information about an image.
    
    Attributes:
        filename (str): filepath to the png file
        caption (str): Caption for the image file
    """
    filename: str
    caption: str