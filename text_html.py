from termcolor import colored
from msgraph import GraphServiceClient


imie=''
nazwisko=''
firma=''

def body(imie, nazwisko, firma):
    body= f"""    
    <p>Szanowny Panie {imie}
    \n 
    {nazwisko} Pragniemy poinformować, że w Domu przy <strong>Staszica</strong> w sprzedaży zostały ostatnie wolne mieszkania.Nasza nowa inwestycja to kameralny budynek położony na południu Krakowa, w bezpośrednim sąsiedztwie prawie dziesięciohektarowego Parku Lilli Wenedy.
    Będzie się w nim znajdować 18 mieszkań, z czego sześć to dwupoziomowe apartamenty z antresolami. Prace budowlane postępują zgodnie z harmonogramem, planowany termin oddania inwestycji do użytkowania to przełom sierpnia i września przyszłego roku.Zachęcamy do zapoznania się z ofertą na naszej stronie internetowej kbi.krakow.pl. Znajdują się tam szczegółowe informacje dotyczące inwestycji oraz wyszukiwarka lokali, która ułatwi znalezienie wymarzonego mieszkania. Serdecznie zapraszamy do rezerwowania mieszkań 
    mailowo lub telefonicznie, a w razie pytań pozostajemy do dyspozycji.Krakowskie Biuro Inwestycyjne
    \n
    Mailing wysłany dla firmy {firma}
    Mailing został zrealizowany przez Krajowy Rynek Nieruchomości sp. z o.o. ul. Ignacego Krasickiego 36A, 30-503 Kraków\n
    \n
    Aby anulować subskrypcję kliknij tutaj</p>   
    """
    return body