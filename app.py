"""
Valente Studio - Servidor local
Ejecutar: python app.py
Abre: http://localhost:5000
"""

import http.server
import socketserver
import os
import json
import urllib.parse
from datetime import datetime

PORT = 5000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "products.json")


def load_products():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_products(products):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(products, f, ensure_ascii=False, indent=2)


class ValenteHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        # API: listar productos
        if parsed.path == "/api/products":
            products = load_products()
            self._json_response(200, products)

        # API: estadísticas
        elif parsed.path == "/api/stats":
            products = load_products()
            total = len(products)
            pending = sum(1 for p in products if not p.get("purchased"))
            purchased = total - pending
            total_amount = sum(
                (p.get("price", 0) or 0) * (p.get("qty", 1) or 1)
                for p in products
            )
            pending_amount = sum(
                (p.get("price", 0) or 0) * (p.get("qty", 1) or 1)
                for p in products if not p.get("purchased")
            )
            self._json_response(200, {
                "total": total,
                "pending": pending,
                "purchased": purchased,
                "total_amount": total_amount,
                "pending_amount": pending_amount,
                "exported_at": datetime.now().isoformat()
            })

        # Archivos estáticos (HTML, CSS, JS)
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)

        # API: guardar producto nuevo
        if parsed.path == "/api/products":
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            product = json.loads(body)
            products = load_products()
            product["id"] = product.get("id") or str(int(datetime.now().timestamp() * 1000))
            product["createdAt"] = product.get("createdAt") or datetime.now().isoformat()
            products.insert(0, product)
            save_products(products)
            self._json_response(201, product)
        else:
            self._json_response(404, {"error": "Ruta no encontrada"})

    def do_PUT(self):
        parsed = urllib.parse.urlparse(self.path)

        # API: actualizar producto — /api/products/<id>
        if parsed.path.startswith("/api/products/"):
            product_id = parsed.path.split("/")[-1]
            length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(length)
            updated = json.loads(body)
            products = load_products()
            products = [updated if p["id"] == product_id else p for p in products]
            save_products(products)
            self._json_response(200, updated)
        else:
            self._json_response(404, {"error": "Ruta no encontrada"})

    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)

        # API: eliminar producto — /api/products/<id>
        if parsed.path.startswith("/api/products/"):
            product_id = parsed.path.split("/")[-1]
            products = load_products()
            products = [p for p in products if p["id"] != product_id]
            save_products(products)
            self._json_response(200, {"deleted": product_id})
        else:
            self._json_response(404, {"error": "Ruta no encontrada"})

    def _json_response(self, status, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {format % args}")


if __name__ == "__main__":
    with socketserver.TCPServer(("", PORT), ValenteHandler) as httpd:
        print(f"\n  Valente Studio corriendo en http://localhost:{PORT}")
        print(f"  Directorio: {BASE_DIR}")
        print(f"  Presiona Ctrl+C para detener\n")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n  Servidor detenido.")
