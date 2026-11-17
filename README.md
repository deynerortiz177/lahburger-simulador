# La H Burger · ¿Cuánto podemos pedir prestado? (Streamlit)

Simulador Monte Carlo del flujo de caja de La H Burger. Responde cuánto puede pedir prestado el negocio,
a qué plazo y para qué destino, sin que el riesgo de no cubrir la cuota supere el límite aceptado.

## Archivos
| Archivo | Para qué sirve |
|---|---|
| `app.py` | Interfaz de Streamlit (colores y logo de la marca) |
| `modelo.py` | Toda la lógica: carga, estadística, Monte Carlo, monto máximo, plazos, destinos y requisitos. Se importa igual en Google Colab |
| `.streamlit/config.toml` | Tema negro y dorado de la marca |
| `datos/Datos_Simulados_LaHBurger.xlsx` | Datos simulados con las reglas del negocio (reemplazar por los reales) |
| `datos/Plantilla_Registro_LaHBurger.xlsx` | Plantilla vacía |
| `datos/logo.png` | Logo |
| `requirements.txt` | Librerías |

## Ejecutar en su computador
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar gratis (Streamlit Community Cloud)
1. Suba esta carpeta completa, con las carpetas `datos` y `.streamlit`, a un repositorio de GitHub.
2. Entre a https://share.streamlit.io y elija **Create app**: repositorio, rama `main`, archivo `app.py`, luego **Deploy**.

## Reproducir en Google Colab
```python
!pip install openpyxl
# suba modelo.py y Datos_Simulados_LaHBurger.xlsx al panel de archivos de Colab
from modelo import *
datos, avisos = cargar_excel("Datos_Simulados_LaHBurger.xlsx")
R = resumir(datos)                          # periodo base desde la hoja Parametros (24/09/2026)
p = datos["Parametros"]
S = Supuestos(cv_min=R.cv_hist["min"], cv_moda=R.cv_hist["moda"], cv_max=R.cv_hist["max"],
              costos_fijos=R.costos_fijos_mes, retiro_min=p["Retiro_min"], retiro_max=p["Retiro_max"],
              choque_mensual=p["Choque_mensual"])
sin_prestamo = simular(R, S)
print(sin_prestamo.prob_f_op_positivo(), sin_prestamo.percentiles())

credito = Credito(monto=0, tasa_ea=p["Tasa_EA"], plazo=int(p["Plazo_meses"]),
                  comision_inicial=p["Comision_inicial"], seguro_mensual=p["Seguro_mensual"])
print(monto_maximo(sin_prestamo, credito))   # monto máximo que el negocio puede pagar
destinos = [Destino.desde_fila(f) for _, f in datos["Escenarios"].iterrows()]
print(comparar_destinos(R, S, destinos[:-1], credito))
print(comparar_plazos(R, S, destinos[0], credito))
```
