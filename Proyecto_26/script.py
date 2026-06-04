import re
from bs4 import BeautifulSoup
from pathlib import Path
import json

def analizar_proyecto(ruta_html, ruta_css, ruta_js):
    # 1. Leer archivos
    html = Path(ruta_html).read_text(encoding='utf-8')
    css = Path(ruta_css).read_text(encoding='utf-8') 
    js = Path(ruta_js).read_text(encoding='utf-8')
    
    soup = BeautifulSoup(html, 'html.parser')
    
    resultado = {
        "css_usado": [],
        "css_huerfano": [],
        "ids_usados_js": [],
        "clases_usadas_js": [],
        "eventos_detectados": []
    }
    
    # 2. Extraer todos los selectores del CSS
    selectores_css = re.findall(r'([.#]?[a-zA-Z][\w-]*)\s*{', css)
    selectores_css = list(set(selectores_css)) # quitar duplicados
    
    # 3. Ver cuáles selectores sí existen en el HTML
    for selector in selectores_css:
        if selector.startswith('.'): # clase
            clase = selector[1:]
            if soup.find(class_=clase):
                resultado["css_usado"].append(selector)
            else:
                resultado["css_huerfano"].append(selector)
        elif selector.startswith('#'): # id
            id_val = selector[1:]
            if soup.find(id=id_val):
                resultado["css_usado"].append(selector)
            else:
                resultado["css_huerfano"].append(selector)
        else: # tag
            if soup.find(selector):
                resultado["css_usado"].append(selector)
            else:
                resultado["css_huerfano"].append(selector)
    
    # 4. Buscar en JS qué IDs y clases está tocando
    ids_en_js = re.findall(r'getElementById\([\'"](.+?)[\'"]\)', js)
    ids_en_js += re.findall(r'querySelector\([\'"]#(.+?)[\'"]\)', js)
    resultado["ids_usados_js"] = list(set(ids_en_js))
    
    clases_en_js = re.findall(r'getElementsByClassName\([\'"](.+?)[\'"]\)', js)
    clases_en_js += re.findall(r'querySelectorAll?\(.*?\.([a-zA-Z][\w-]*)', js)
    resultado["clases_usadas_js"] = list(set(clases_en_js))
    
    # 5. Detectar eventos básicos
    eventos = re.findall(r'addEventListener\([\'"](\w+)[\'"]', js)
    eventos += re.findall(r'\.on(\w+)\s*=', js)
    resultado["eventos_detectados"] = list(set(eventos))
    
    return resultado

# Uso:
if __name__ == "__main__":
    reporte = analizar_proyecto("index.html", "styles.css", "script.js")
    print(json.dumps(reporte, indent=2, ensure_ascii=False))