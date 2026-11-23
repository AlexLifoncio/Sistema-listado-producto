from dataclasses import dataclass
from datetime import date

@dataclass
class Producto:
    idProducto : int | str
    producto : str
    fecha : date
    cantidad : int | float
    precio : float
    subtotal : float
