import io
import string
from typing import Any, Dict, List

from PIL import Image
from openpyxl import load_workbook


class SheetImageLoader:
    """Loads all images in a sheet"""

    def __init__(self, sheet):
        """Loads all sheet images"""
        
        self._images:Dict[str, Any] = {}
        
        sheet_images = sheet._images
        for image in sheet_images:
            row = image.anchor._from.row + 1
            col = string.ascii_uppercase[image.anchor._from.col]
            self._images[f'{col}{row}'] = image._data

    def image_in(self, cell):
        """Checks if there's an image in specified cell"""
        return cell in self._images

    def get_image_anchors(self) -> List[str]:
        return list(self._images.keys())

    def get(self, cell):
        """Retrieves image data from a cell"""
        if cell not in self._images:
            raise ValueError("Cell {} doesn't contain an image".format(cell))
        else:
            image = io.BytesIO(self._images[cell]())
            return Image.open(image)
        
if __name__ == "__main__":
    
    wb = load_workbook(filename="assets/fiche.xlsx", data_only=True)
    ws = wb["Suivi"]
    
    image_loader = SheetImageLoader(ws)
    image = image_loader.get("B43")