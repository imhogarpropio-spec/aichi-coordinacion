# --- CONFIG (ajústalo a tu entorno) ---
$Env:PGHOST = $Env:PGHOST  # p.ej. "localhost"
$Env:PGPORT = $Env:PGPORT  # p.ej. "5432"
$Env:PGUSER = $Env:PGUSER  # p.ej. "postgres"
$Env:PGPASSWORD = $Env:Imenes137323$  # tu password
$DB_NAME_DELEG = "delegacionesdb"    # <- NOMBRE de la BD de delegaciones (no la de usuarios)

Write-Host ">> Reiniciando solo la BD '$DB_NAME_DELEG'..."
psql -U $Env:PGUSER -h $Env:PGHOST -p $Env:PGPORT -d postgres -c "SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname='$DB_NAME_DELEG' AND pid <> pg_backend_pid();" | Out-Null
psql -U $Env:PGUSER -h $Env:PGHOST -p $Env:PGPORT -d postgres -c "DROP DATABASE IF EXISTS $DB_NAME_DELEG;" | Out-Null
psql -U $Env:PGUSER -h $Env:PGHOST -p $Env:PGPORT -d postgres -c "CREATE DATABASE $DB_NAME_DELEG;" | Out-Null

Write-Host ">> Creando SOLO tablas de delegaciones (no 'usuarios')..."
$Env:FLASK_APP = "app:create_app"
flask shell -c "
from flask import current_app as app
from extensiones import db
from models import *
# 1) crear tablas del bind por DEFECTO (delegaciones)
db.create_all(app=app)
# 2) crear tablas de binds EXCEPTO 'usuarios'
binds = (app.config.get('SQLALCHEMY_BINDS') or {}).keys()
for b in binds:
    if b != 'usuarios':
        db.create_all(app=app, bind=b)
print('OK: Tablas creadas para default y binds != usuarios')
" | Out-Null

Write-Host ">> Listo. 'usuariosdb' no fue tocada ni se creó ningún usuario nuevo."
