# LPN Matcher - Chedraui (app web)

App Flask que envuelve el script `lpn_matcher.py` original: subes el Excel de
pedido y el PDF de etiquetas, y descargas un Excel con la columna LPN
agregada.

## Correr en local (requiere Docker, por poppler/tesseract)

```
docker build -t lpn-matcher .
docker run -p 5000:5000 lpn-matcher
```

Abre http://localhost:5000

Sin Docker, en Linux/macOS puedes instalar las dependencias del sistema y
correr directo:

```
sudo apt-get install poppler-utils tesseract-ocr tesseract-ocr-spa libzbar0
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

En Windows no hay binarios oficiales sencillos de poppler/tesseract para
`pdftoppm`/`tesseract` en PATH, por eso se recomienda Docker incluso para
probar en local.

## Desplegar en linea

Cualquier plataforma que soporte "Deploy from Dockerfile" sirve, por ejemplo:

**Render**
1. Sube este proyecto a un repo de GitHub.
2. En Render: New -> Web Service -> selecciona el repo.
3. Render detecta el `Dockerfile` automaticamente. Deja el puerto en 5000
   (o usa la variable `PORT` que Render inyecta, ya soportada en el Dockerfile).
4. Define la variable de entorno `SECRET_KEY` con un valor aleatorio.
5. Deploy.

**Railway / Fly.io**: mismo flujo, "Deploy from Dockerfile", detectan el
`Dockerfile` y exponen el puerto automaticamente.

## Notas

- Los archivos subidos se procesan en una carpeta temporal y se borran al
  terminar. El Excel de salida se genera en `outputs/` y se borra despues de
  que el usuario lo descarga.
- Limite de subida: 40 MB por archivo (ajustable en `app.py`,
  `MAX_CONTENT_LENGTH`).
- El OCR puede tardar varios minutos con PDFs de muchas paginas; en
  produccion considera subir el timeout del servidor/proxy si el pedido es
  grande (el `Dockerfile` ya arranca gunicorn con `-t 300`).

## Catalogo maestro (catalogo_chedraui.json)

`catalogo_chedraui.json` es el catalogo de productos de Chedraui/AKSI (UPC,
Codigo, SKU de Chedraui, Descripcion), independiente de cualquier pedido. Se
usa como respaldo en `match()` (matcher.py): si el OCR lee mal un digito del
UPC impreso en el PDF y no hay match exacto, se busca el UPC valido mas
parecido en este catalogo (solo si hay un unico candidato a 1 digito de
distancia, sin ambiguedad) para no caer directo a la comparacion difusa de
descripcion.

Para actualizarlo cuando cambie el catalogo de productos, desde un Excel con
columnas SKU, ARTICULO, #, AKSI (=Codigo), ARTICULO, UPC, Descripcion:

```python
import openpyxl, json, re
wb = openpyxl.load_workbook("nuevo_catalogo.xlsx", data_only=True)
ws = wb.worksheets[0]
catalogo, vistos = [], set()
for row in ws.iter_rows(min_row=2, values_only=True):
    sku, upc, aksi, desc = row[0], row[5], row[3], row[6]
    if aksi is None or upc is None:
        continue
    aksi_str = str(aksi).strip()
    if aksi_str in vistos:
        continue
    vistos.add(aksi_str)
    catalogo.append({
        "upc": re.sub(r"\D", "", str(upc)),
        "codigo": aksi_str,
        "sku_chedraui": str(sku).strip(),
        "descripcion": (desc or "").strip(),
    })
json.dump(catalogo, open("catalogo_chedraui.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
```
