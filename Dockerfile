# استخدام Python 3.10 - نسخة مستقرة ومتوافقة 100%
FROM python:3.10-slim-bullseye

# تعيين متغيرات البيئة
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

# تثبيت المكتبات الأساسية للنظام
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    g++ \
    gfortran \
    libopenblas-dev \
    liblapack-dev \
    && rm -rf /var/lib/apt/lists/*

# إنشاء مجلد التطبيق
WORKDIR /app

# نسخ ملفات المتطلبات أولاً (للاستفادة من Docker cache)
COPY requirements.txt .

# ترقية pip وتثبيت المكتبات
RUN pip install --upgrade pip setuptools wheel && \
    pip install --no-cache-dir -r requirements.txt

# نسخ الكود
COPY app.py .

# إنشاء مستخدم غير root للأمان
RUN useradd -m -u 1000 appuser && chown -R appuser:appuser /app
USER appuser

# فتح المنفذ
EXPOSE 5000

# تشغيل التطبيق
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--threads", "4", "--timeout", "120", "app:app"]
