# =========================================================
# OCR - DETECTA PLACA (letras grandes con guion) / SERIE
# =========================================================

LETRAS_A_DIGITO = {"O": "0", "Q": "0", "D": "0", "I": "1", "L": "1",
                   "Z": "2", "S": "5", "B": "8", "G": "6"}
DIGITO_A_LETRA = {"0": "O", "1": "I", "2": "Z", "5": "S", "8": "B", "6": "G"}

VIN_VALIDO = re.compile(r"^[A-HJ-NPR-Z0-9]{17}$")

# Formatos de placa (L = letra, D = dígito), sin guiones
#   R20-BSD   -> LDD-LLL      (CDMX)
#   MKJ-121-B -> LLL-DDD-L    (Edomex)
ESQUEMAS_PLACA = ["LDDLLL", "LLLDDDL"]

# Letras permitidas por Tesseract (ahora SÍ incluye el guion)
CONFIG_OCR = (
    "-c tessedit_char_whitelist="
    "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789- "
)


def preparar_imagen(img, angulo=0, invertir=False, lado_objetivo=2000):

    img = img.convert("L")

    if angulo:
        img = img.rotate(angulo, expand=True)

    lado = max(img.size)
    factor = lado_objetivo / lado
    img = img.resize(
        (int(img.width * factor), int(img.height * factor)),
        Image.LANCZOS
    )

    img = ImageOps.autocontrast(img, cutoff=2)

    if invertir:
        img = ImageOps.invert(img)

    return img


def corregir_con_esquema(ventana, esquema):
    """Corrige O/0, I/1, S/5, B/8... según si ahí va letra o dígito."""

    if len(ventana) != len(esquema):
        return None

    resultado = ""

    for ch, tipo in zip(ventana, esquema):

        if tipo == "D":
            ch = LETRAS_A_DIGITO.get(ch, ch)
            if not ch.isdigit():
                return None
        else:
            ch = DIGITO_A_LETRA.get(ch, ch)
            if not ch.isalpha():
                return None

        resultado += ch

    return resultado


def placas_en_linea(texto_linea):
    """Busca placas dentro de una línea. Devuelve placas SIN guiones."""

    limpio = re.sub(r"[^A-Z0-9]", "", texto_linea.upper())
    encontradas = []

    for esquema in ESQUEMAS_PLACA:
        n = len(esquema)
        for i in range(len(limpio) - n + 1):
            p = corregir_con_esquema(limpio[i:i + n], esquema)
            if p:
                encontradas.append(p)

    return encontradas


def placas_texto_grande(img_preparada, psm):
    """
    Lee la imagen con posiciones y se queda con el texto de letras GRANDES.
    Devuelve lista de (placa_sin_guiones, puntaje).
    """

    data = pytesseract.image_to_data(
        img_preparada,
        config=f"--psm {psm} {CONFIG_OCR}",
        output_type=pytesseract.Output.DICT
    )

    palabras = []

    for i, t in enumerate(data["text"]):

        t = str(t).strip().upper()

        if not t or not re.search(r"[A-Z0-9\-]", t):
            continue

        try:
            conf = float(data["conf"][i])
        except Exception:
            conf = 0

        if conf < 10:
            continue

        palabras.append({
            "t": t,
            "x": data["left"][i],
            "y": data["top"][i],
            "h": data["height"][i],
        })

    if not palabras:
        return []

    # Agrupar palabras en renglones (mismo centro vertical aprox.)
    palabras.sort(key=lambda p: p["y"] + p["h"] / 2)

    renglones = []

    for p in palabras:
        centro = p["y"] + p["h"] / 2
        if renglones:
            ult = renglones[-1]
            if abs(centro - ult["centro"]) < ult["h"] * 0.6:
                ult["items"].append(p)
                continue
        renglones.append({"centro": centro, "h": p["h"], "items": [p]})

    h_max = max(p["h"] for p in palabras)

    resultados = []

    for r in renglones:

        items = sorted(r["items"], key=lambda p: p["x"])
        texto = " ".join(p["t"] for p in items)
        altura = max(p["h"] for p in items)

        # Ignorar letras chicas (leyendas, municipio, etc.)
        if altura < h_max * 0.5:
            continue

        tiene_guion = "-" in texto

        for placa in placas_en_linea(texto):
            puntos = 10 + 20 * (altura / h_max) + (8 if tiene_guion else 0)
            resultados.append((placa, puntos))

    return resultados


def extraer_vin(texto):
    """Busca números de serie (VIN de 17 caracteres) en el texto."""

    candidatos = []

    texto = texto.upper().replace("*", " ").replace("-", " ")

    for linea in texto.splitlines():

        limpio = re.sub(r"[^A-Z0-9]", "", linea)

        if len(limpio) < 17:
            continue

        base_vin = limpio
        extra = 0

        if "VIN" in limpio:
            base_vin = limpio.split("VIN", 1)[1]
            extra = 10

        for i in range(len(base_vin) - 16):

            v = base_vin[i:i + 17]
            v = v.replace("O", "0").replace("Q", "0").replace("I", "1")

            if (
                VIN_VALIDO.match(v)
                and any(c.isdigit() for c in v)
                and any(c.isalpha() for c in v)
            ):
                candidatos.append((v, 20 + extra))

    return candidatos


def leer_texto_foto(imagen_bytes, claves_conocidas=None):

    try:

        img = Image.open(io.BytesIO(imagen_bytes))
        img = ImageOps.exif_transpose(img).convert("RGB")

        conocidas = set(claves_conocidas or [])
        puntajes = {}

        for angulo in (0, 180, 90, 270):

            for invertir in (False, True):

                for lado in (2000, 3000):

                    trabajo = preparar_imagen(img, angulo, invertir, lado)

                    for psm in (11, 6):

                        # 1) Placa: solo letras grandes, sin guiones
                        try:
                            for cand, pts in placas_texto_grande(trabajo, psm):
                                puntajes[cand] = puntajes.get(cand, 0) + pts
                        except Exception:
                            pass

                        # 2) Serie / VIN
                        texto = pytesseract.image_to_string(
                            trabajo,
                            config=f"--psm {psm} {CONFIG_OCR}"
                        )
                        for cand, pts in extraer_vin(texto):
                            puntajes[cand] = puntajes.get(cand, 0) + pts

            if puntajes and max(puntajes.values()) >= 40:
                break

        if not puntajes:
            return ""

        # 1) Lo que exista en el Excel gana
        en_base = [c for c in puntajes if c in conocidas]
        if en_base:
            return max(en_base, key=lambda c: puntajes[c])

        # 2) Coincidencia aproximada (si falla 1 carácter)
        if conocidas:

            lista = list(conocidas)

            for cand in sorted(puntajes, key=puntajes.get, reverse=True):
                corte = 0.93 if len(cand) == 17 else 0.83
                cerca = get_close_matches(cand, lista, n=1, cutoff=corte)
                if cerca:
                    return cerca[0]

        # 3) Mejor candidato por puntaje
        return max(puntajes, key=puntajes.get)

    except Exception:
        return ""
