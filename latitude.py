from geopy.geocoders import ArcGIS


def obter_lat_long(endereco):
    geolocalizador = ArcGIS()
    localizacao = geolocalizador.geocode(endereco)

    if localizacao:
        latitude = localizacao.latitude
        longitude = localizacao.longitude
        return latitude, longitude
    else:
        return None


# Exemplo de uso
endereco = "R. S-1, Q. 153 - L. 25 - St. Bueno, Goiânia - GO, 74230-220"
coordenadas = obter_lat_long(endereco)

if coordenadas:
    print(f"Latitude: {coordenadas[0]}, Longitude: {coordenadas[1]}")
else:
    print("Endereço não encontrado.")