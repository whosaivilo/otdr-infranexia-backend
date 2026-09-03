#!/bin/bash
php artisan config:cache
php artisan route:cache
# Matikan baris di bawah ini kalau belum ada database
# php artisan migrate --force

php-fpm -D
nginx -g "daemon off;"
