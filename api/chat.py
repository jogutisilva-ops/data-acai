from http.server import BaseHTTPRequestHandler
import json
import os
import urllib.request

# Get path of data_summary.json relative to this file
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
data_file_path = os.path.join(parent_dir, 'data_summary.json')

try:
    with open(data_file_path, 'r', encoding='utf-8') as f:
        data_summary = f.read()
except Exception as e:
    data_summary = "{}"

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        # Read request body
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            req_data = json.loads(body.decode('utf-8'))
            user_message = req_data.get('message', '')
        except Exception as e:
            self.send_response(400)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': 'Invalid JSON body: ' + str(e)}).encode('utf-8'))
            return

        # Get API Key
        api_key = os.environ.get('GEMINI_API_KEY')
        if not api_key:
            self.send_response(200)  # Return 200 with an error description to show nicely in chat fallback
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error_no_key': True, 'response': 'La clave de API (GEMINI_API_KEY) no está configurada en el panel de Vercel.'}).encode('utf-8'))
            return

        api_key = api_key.strip()

        # Call Gemini API using standard library urllib
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent"
        
        system_instruction = (
            "Eres un asistente virtual experto en análisis de datos comerciales de Açaí Prime.\n"
            "Tienes acceso al siguiente resumen estructurado de ventas, comisiones y consolidación del periodo (11 de febrero al 25 de mayo de 2026):\n\n"
            f"{data_summary}\n\n"
            "Reglas:\n"
            "1. Responde de forma concisa, profesional y amigable en español.\n"
            "2. Basa tus respuestas matemáticas estrictamente en los datos numéricos provistos en el resumen (revisa KPIs, productos, métodos de pago, ventas semanales, etc.).\n"
            "3. Explica los criterios de consolidación aplicados para bebidas, cafés, jugos y aguas cuando sea pertinente.\n"
            "4. Si te preguntan sobre cosas que no están en el resumen, indícalo de forma cortés."
        )

        gemini_req_body = {
            "contents": [
                {
                    "role": "user",
                    "parts": [{"text": user_message}]
                }
            ],
            "systemInstruction": {
                "parts": [{"text": system_instruction}]
            },
            "generationConfig": {
                "temperature": 0.2
            }
        }

        req = urllib.request.Request(
            url,
            data=json.dumps(gemini_req_body).encode('utf-8'),
            headers={
                'Content-Type': 'application/json',
                'X-goog-api-key': api_key
            },
            method='POST'
        )

        try:
            with urllib.request.urlopen(req) as response:
                res_body = response.read().decode('utf-8')
                res_data = json.loads(res_body)
                
                # Extract text response from Gemini format
                text_response = ""
                if 'candidates' in res_data and len(res_data['candidates']) > 0:
                    candidate = res_data['candidates'][0]
                    if 'content' in candidate and 'parts' in candidate['content']:
                        parts = candidate['content']['parts']
                        if len(parts) > 0 and 'text' in parts[0]:
                            text_response = parts[0]['text']
                
                if not text_response:
                    text_response = "No se recibió respuesta del modelo de lenguaje."

                self.send_response(200)
                self.send_header('Content-Type', 'application/json')
                self.send_header('Access-Control-Allow-Origin', '*')
                self.end_headers()
                self.wfile.write(json.dumps({'response': text_response}).encode('utf-8'))
        except Exception as e:
            key_preview = api_key[:6] + "..." if api_key else "None"
            key_len = len(api_key) if api_key else 0
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Access-Control-Allow-Origin', '*')
            self.end_headers()
            self.wfile.write(json.dumps({'error': f'Error de conexión con Gemini (Key: {key_preview}, len: {key_len}): {str(e)}'}).encode('utf-8'))
