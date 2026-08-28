# syntax=docker/dockerfile:1
# ============================================================================
#  HRM - Phan mem quan ly nhan su (Flask + SQL Server)
#  Python 3.12 / Debian 12 bookworm / ODBC Driver 17 for SQL Server
# ============================================================================
FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONIOENCODING=utf-8 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    LANG=C.UTF-8 \
    TZ=Asia/Ho_Chi_Minh

# --- 1. He thong + Microsoft ODBC Driver 17 for SQL Server -----------------
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
        ca-certificates curl gnupg apt-transport-https unixodbc tzdata; \
    curl -fsSL https://packages.microsoft.com/keys/microsoft.asc \
        | gpg --dearmor -o /usr/share/keyrings/microsoft-prod.gpg; \
    echo "deb [arch=amd64,arm64 signed-by=/usr/share/keyrings/microsoft-prod.gpg] https://packages.microsoft.com/debian/12/prod bookworm main" \
        > /etc/apt/sources.list.d/mssql-release.list; \
    apt-get update; \
    ACCEPT_EULA=Y apt-get install -y --no-install-recommends msodbcsql17; \
    rm -rf /var/lib/apt/lists/*

# --- 2. Alias driver "SQL Server" + Threading ------------------------------
#     a) unixODBC mac dinh serialize TOAN BO loi goi ODBC neu driver khong khai
#        bao "Threading". App chay nhieu thread (waitress/uvicorn) va mo
#        connection moi cho tung request se bi xep hang -> "Login timeout expired".
#        Dat Threading=1 = chi khoa theo tung connection.
#     b) Code co cho dung DRIVER={SQL Server} (ten cua Windows, Linux khong co)
#        -> nhan ban nguyen section thanh alias, KHONG phai chi moi dong Driver=.
RUN set -eux; \
    INI=/etc/odbcinst.ini; \
    SRC='ODBC Driver 17 for SQL Server'; \
    grep -q '^Threading' "$INI" || sed -i "/^\[$SRC\]/a Threading=1" "$INI"; \
    awk -v s="[$SRC]" 'BEGIN{p=0}{ if($0==s){p=1; print "[SQL Server]"; next} if(p && /^\[/){p=0} if(p) print }' "$INI" > /tmp/alias.ini; \
    odbcinst -i -d -f /tmp/alias.ini; \
    rm -f /tmp/alias.ini; \
    cat "$INI"

# --- 3. OpenSSL 3: cho phep TLS cu + chu ky SHA1 (SQL Server doi cu) -------
#     Khong co buoc nay se loi:
#       SSL Provider: [error:0A00014D:SSL routines::legacy sigalg disallowed]
COPY fix-openssl-legacy.py /tmp/fix-openssl-legacy.py
RUN python /tmp/fix-openssl-legacy.py && rm -f /tmp/fix-openssl-legacy.py

WORKDIR /app

# --- 4. Thu vien Python ----------------------------------------------------
#     PIP_INDEX_URL: doi mirror khi mang cham (xem .env / README muc 13)
ARG PIP_INDEX_URL=https://pypi.org/simple
COPY requirements.docker.txt ./requirements.docker.txt
RUN --mount=type=cache,target=/root/.cache/pip,id=nt-pip,sharing=locked \
    pip install --retries 10 --timeout 60 \
        --index-url "$PIP_INDEX_URL" \
        -r requirements.docker.txt

# --- 5. Source code --------------------------------------------------------
COPY . /app
RUN mkdir -p /app/nhapxuat/nhap /app/nhapxuat/xuat /app/nhapxuat/bienban \
             /app/static/uploads/mau /app/static/img/avatar /app/static/img/qr

EXPOSE 81
HEALTHCHECK --interval=30s --timeout=10s --start-period=40s --retries=3 \
    CMD curl -fsS -o /dev/null http://127.0.0.1:81/ || exit 1

# Production: waitress (giong product.bat)
CMD ["python", "main.py"]
