# 1. Gunakan server dasar PHP 8.2 (sesuaikan dengan versi PHP kalian)
FROM php:8.2-fpm

# 2. Install dependensi sistem, web server (Nginx), dan Python
RUN apt-get update && apt-get install -y \
    nginx \
    zip \
    unzip \
    git \
    python3 \
    python3-pip \
    python3-venv

# 3. Buat Virtual Environment Python & Install library untuk Excel
RUN python3 -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"
RUN pip install pandas openpyxl

# 4. Install ekstensi PHP untuk Database (MySQL & PostgreSQL)
RUN docker-php-ext-install pdo_mysql pdo_pgsql

# 5. Install Composer untuk dependensi Laravel
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# 6. Tentukan folder kerja
WORKDIR /var/www/html

# 7. Copy seluruh file project Laravel kalian ke dalam server
COPY . .

# 8. Install library Laravel
RUN composer install --no-dev --optimize-autoloader

# 9. Berikan izin akses folder agar Laravel bisa upload/simpan cache
RUN chown -R www-data:www-data /var/www/html/storage /var/www/html/bootstrap/cache

# 10. Copy settingan Nginx dan script jalankan server
COPY nginx.conf /etc/nginx/sites-available/default
COPY start.sh /start.sh
RUN chmod +x /start.sh

# Buka port 80 untuk Render
EXPOSE 80

# Perintah yang dijalankan saat server menyala
CMD ["/start.sh"]
