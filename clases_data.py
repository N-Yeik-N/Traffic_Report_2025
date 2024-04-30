from dataclasses import dataclass, asdict
from datetime import datetime

@dataclass
class folderVariables:
    numsubarea: str
    nameproject: str
    nomdistrito: str
    #presinter: str

@dataclass
class typicalInterseccion:
    codinterseccion: str
    nominterseccion: str
    hpinterseccionmt: str
    hpintersecciontt: str
    hpinterseccionnt: str

@dataclass
class atypicalInterseccion:
    codinterseccion: str
    nominterseccion: str
    hpinterseccionma: str
    hpinterseccionta: str
    hpinterseccionna: str

@dataclass
class atypicalSystem:
    hpsistemama: str
    hpsistemata: str
    hpsistemana: str

@dataclass
class typicalSystem:
    hpsistemamt: str
    hpsistematt: str
    hpsistemant: str

@dataclass #No lo use
class typicalTales:
    codigos: list
    dcolast: str

@dataclass #No lo use
class atypicalTales:
    codigos: list
    dcolasa: str