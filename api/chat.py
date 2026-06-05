from http.server import BaseHTTPRequestHandler
import json
import os
import sqlite3
import urllib.request
import re

# Get path of sales_data.db relative to this file
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
db_file_path = os.path.join(parent_dir, 'sales_data.db')

system_instruction = (
    "Eres un asistente virtual experto en análisis de datos comerciales de Açaí Prime.\n"
    "Tienes acceso a la base de datos de transacciones de ventas completa (tabla 'ventas') del periodo (11 de febrero al 25 de mayo de 2026).\n\n"
    "Estructura de la tabla 'ventas':\n"
    "- id_transaccion (TEXT): ID de la transacción.\n"
    "- fecha_original (TEXT): Fecha en formato original 'DD-MM-YYYY, HH:MM'.\n"
    "- fecha_datetime (TEXT): Fecha formateada como 'YYYY-MM-DD HH:MM:SS' (apta para filtros de fecha en SQLite, por ejemplo, strftime('%H', fecha_datetime)).\n"
    "- semana (TEXT): Semana en formato 'YYYY-MM-DD/YYYY-MM-DD' (ej: '2026-02-16/2026-02-22').\n"
    "- dia_semana (TEXT): Día de la semana (Lunes, Martes, Miércoles, Jueves, Viernes, Sábado, Domingo).\n"
    "- descripcion_original (TEXT): Descripción original del ítem en el Excel.\n"
    "- producto_limpio (TEXT): Nombre del producto unificado y limpio (ej: 'AÇAÍ SMALL (260 ML) - AÇAÍ CLÁSICO', 'CHEESE BURGER').\n"
    "- categoria (TEXT): Categoría unificada ('AÇAÍ PRIME', 'AMERICAN PRIME BURGER', 'Propina', 'Extras', 'Importe personalizado', 'Otros / Sin Categoría').\n"
    "- cantidad (INTEGER): Cantidad de unidades vendidas.\n"
    "- precio_sin_descuento (REAL): Precio original del producto antes de aplicar descuentos en CLP.\n"
    "- descuento (REAL): Monto total del descuento aplicado al producto en CLP (precio_sin_descuento - precio_bruto).\n"
    "- precio_bruto (REAL): Ingreso bruto en CLP (precio final pagado por el cliente after discount).\n"
    "- precio_neto (REAL): Ingreso neto en CLP.\n"
    "- forma_pago (TEXT): Método de pago (ej: 'Tarjeta de Débito', 'Tarjeta de Crédito', 'Efectivo', 'No especificado').\n"
    "- fee_rate (REAL): Comisión cobrada (ej: 0.0155 para débito, 0.025 para crédito/Visa/Amex, 0.00 para efectivo).\n"
    "- fee_amount (REAL): Comisión cobrada en CLP (precio_bruto * fee_rate).\n"
    "- net_after_fee (REAL): Ingreso real neto tras comisiones (precio_bruto - fee_amount).\n"
    "- sku (TEXT): Código SKU único del producto.\n"
    "- dispositivo (TEXT): Número de serie del terminal que realizó la venta.\n\n"
    "Reglas:\n"
    "1. Si el usuario te hace una pregunta que requiera consultar datos, debes llamar a la función `run_sql_query` escribiendo una consulta SQL SELECT válida sobre la tabla 'ventas'.\n"
    "2. Escribe consultas SQL limpias y optimizadas. Si necesitas agrupar por producto o semana, usa `GROUP BY` y ordena con `ORDER BY` de mayor a menor.\n"
    "3. IMPORTANTE: En SQLite las comparaciones de texto con LIKE o de igualdad son case-insensitive. Utiliza `producto_limpio` para nombres de productos y `categoria` para categorías.\n"
    "4. Cuando recibas los datos de la base de datos, presenta la información al usuario en un formato amigable, profesional y estructurado usando Markdown. Redondea los números cuando sea pertinente (ej. comisiones o ingresos) para facilitar la lectura.\n"
    "5. CAPACIDAD DE PROYECCIÓN Y EXTRAPOLACIÓN:\n"
    "   - Si el usuario te solicita proyecciones futuras (ej. 'proyectar ventas del próximo mes' o 'proyección para las siguientes 4 semanas'), NO te niegues.\n"
    "   - Utiliza la data histórica ejecutando consultas SQL para calcular el promedio de venta diario/semanal reciente o la tasa de crecimiento promedio.\n"
    "   - Realiza la extrapolación matemática lineal de forma lógica en tu proceso de razonamiento.\n"
    "   - Explica de manera transparente tu metodología e indica que es una estimación estadística lineal basada en el historial.\n"
    "6. REGLAS DE DESCUENTOS:\n"
    "   - Las ventas con descuento se identifican con `descuento > 0`.\n"
    "   - Para calcular el porcentaje de descuento de un ítem, utiliza la fórmula: `(descuento / precio_sin_descuento) * 100`.\n"
    "   - Para contar cuántas transacciones o boletas tuvieron descuento, usa `COUNT(DISTINCT id_transaccion) FROM ventas WHERE descuento > 0`."
)


tools = [
    {
        "functionDeclarations": [
            {
                "name": "run_sql_query",
                "description": "Ejecuta una consulta SQL SELECT sobre la base de datos de ventas ('ventas') del negocio para obtener métricas, sumatorias, rankings o detalles específicos.",
                "parameters": {
                    "type": "OBJECT",
                    "properties": {
                        "query": {
                            "type": "STRING",
                            "description": "La consulta SQL SELECT a ejecutar. Debe ser válida para SQLite."
                        }
                    },
                    "required": ["query"]
                }
            }
        ]
    }
]

def execute_sql(query):
    # Security checks
    q = query.strip().upper()
    if not q.startswith("SELECT"):
        return {"error": "Solo se permiten consultas SELECT de lectura."}
        
    forbidden = ["DROP", "INSERT", "UPDATE", "DELETE", "ALTER", "CREATE", "REPLACE", "RENAME", "TRUNCATE"]
    for word in forbidden:
        if re.search(r'\b' + word + r'\b', q):
            return {"error": f"Operación no permitida: {word}"}
            
    if not os.path.exists(db_file_path):
        return {"error": f"Base de datos de ventas no encontrada en el servidor backend."}
        
    try:
        conn = sqlite3.connect(db_file_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query)
        # Limit rows to 100 to prevent payload issues
        rows = cursor.fetchmany(100)
        conn.close()
        
        result_list = [dict(row) for row in rows]
        return {"result": result_list}
    except Exception as e:
        return {"error": str(e)}

def call_gemini(payload, api_key):
    import time
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    
    max_retries = 1
    base_delay = 3.0  # seconds
    
    for attempt in range(max_retries + 1):
        req = urllib.request.Request(
            url,
            data=json.dumps(payload).encode('utf-8'),
            headers={
                'Content-Type': 'application/json',
                'X-goog-api-key': api_key
            },
            method='POST'
        )
        try:
            with urllib.request.urlopen(req) as response:
                return json.loads(response.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < max_retries:
                delay = base_delay * (2 ** attempt)
                time.sleep(delay)
                continue
            raise e

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
 
    def do_POST(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            req_data = json.loads(body.decode('utf-8'))
            user_message = req_data.get('message', '')
            client_time = req_data.get('client_time', '')
        except Exception as e:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'Invalid JSON body: ' + str(e)}).encode('utf-8'))
            return

        if not client_time:
            from datetime import datetime
            client_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')

        dynamic_system_instruction = (
            f"FECHA Y HORA ACTUAL DEL CLIENTE (Úsala como referencia absoluta de hoy, ayer, etc. y para contestar preguntas sobre el día actual): {client_time}\n\n"
            f"{system_instruction}"
        )

 
        # Get API Key
        api_key = os.environ.get('GEMINI_API_KEY')
        if not api_key:
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({
                'error_no_key': True, 
                'response': 'La clave de API (GEMINI_API_KEY) no está configurada en el panel de Vercel.'
            }).encode('utf-8'))
            return
 
        api_key = api_key.strip()
 
        # Multi-turn Agent loop to support sequential tool executions (e.g. detailed product breakdowns and counts)
        contents = [
            {
                "role": "user",
                "parts": [{"text": user_message}]
            }
        ]
        
        loop_count = 0
        max_loops = 5
        last_query_executed = None
        final_text = ""
        
        try:
            while loop_count < max_loops:
                gemini_payload = {
                    "contents": contents,
                    "systemInstruction": {
                        "parts": [{"text": dynamic_system_instruction}]
                    },
                    "tools": tools,
                    "generationConfig": {
                        "temperature": 0.0 if loop_count == 0 else 0.2
                    }
                }
                
                res_data = call_gemini(gemini_payload, api_key)
                
                candidate = res_data.get('candidates', [{}])[0]
                content = candidate.get('content', {})
                parts = content.get('parts', [])
                
                # Guardar respuesta del modelo en el historial de conversación
                contents.append(content)
                
                # Buscar si hay llamado a función en la respuesta
                function_call = None
                for part in parts:
                    if 'functionCall' in part:
                        function_call = part['functionCall']
                        break
                        
                if not function_call:
                    # Sin llamadas a función, tenemos la respuesta de texto final de Gemini
                    for part in parts:
                        if 'text' in part:
                            final_text += part['text']
                    break
                    
                func_name = function_call.get('name')
                args = function_call.get('args', {})
                sql_query = args.get('query')
                
                # Guardar la última query ejecutada
                last_query_executed = sql_query
                
                # Ejecutar consulta de lectura SQL
                if func_name == "run_sql_query":
                    query_result = execute_sql(sql_query)
                else:
                    query_result = {"error": f"Función {func_name} no disponible."}
                    
                # Guardar el resultado de la función en el historial
                fc_id = function_call.get('id')
                fn_response_part = {
                    "functionResponse": {
                        "name": func_name,
                        "response": query_result
                    }
                }
                if fc_id:
                    fn_response_part["functionResponse"]["id"] = fc_id
                    
                contents.append({
                    "role": "function",
                    "parts": [fn_response_part]
                })
                
                loop_count += 1
                
            if not final_text:
                final_text = "No se recibió una síntesis de respuesta del modelo."
                
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            
            response_payload = {
                'response': final_text
            }
            if last_query_executed:
                response_payload['query_executed'] = last_query_executed
                
            self.wfile.write(json.dumps(response_payload).encode('utf-8'))
            
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            key_mask = f"{api_key[:6]}...{api_key[-4:]}" if api_key else "None"
            self.wfile.write(json.dumps({'error': f'Error de ejecución en el Agente de Datos: {str(e)} (API Key: {key_mask})'}).encode('utf-8'))

