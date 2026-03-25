# Guía de Deploy

## Opción 1: cPanel con Python App (Passenger)

1. En cPanel → "Setup Python App"
2. Python version: 3.9+
3. Application root: `excelfoborana`
4. Application URL: tu dominio o subdirectorio
5. Application startup file: `wsgi.py`
6. Application Entry point: `application`
7. Guardar → "Run pip install" con el `requirements.txt`

---

## Opción 2: VPS con Gunicorn + Nginx

```bash
# Instalar dependencias
pip install -r requirements.txt

# Correr con gunicorn
gunicorn --bind 0.0.0.0:8000 wsgi:application --workers 2
```

Nginx config (bloque server):
```nginx
location / {
    proxy_pass http://127.0.0.1:8000;
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    client_max_body_size 10M;
}
```

---

## Desarrollo local

```bash
pip install -r requirements.txt
python app.py
# Abrir http://localhost:5000
```
