# Rangos horizontales relativos (normalizados sobre el ancho de página)
x_rel = x_palabra / ancho_pagina

if x_rel < 0.08:
    col_bandeja = palabra
elif x_rel < 0.12:
    col_numero = palabra
elif x_rel < 0.22:
    col_ean = palabra
elif x_rel < 0.36:
    col_nombre.append(palabra)
elif x_rel < 0.46:
    col_marca.append(palabra)
elif x_rel < 0.58:
    col_desc_a.append(palabra)
elif x_rel < 0.67:
    col_fabricante.append(palabra)
elif x_rel < 0.72:
    col_caras = palabra
elif x_rel < 0.79:
    col_altura = palabra
elif x_rel < 0.86:
    col_profundidad = palabra
elif x_rel < 0.94:
    col_total_bandeja = palabra
else:
    col_total_unidades = palabra
