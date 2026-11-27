<?php

use App\Http\Controllers\Api\AuthController;
use App\Http\Controllers\Api\OtdrController;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;

Route::post('/login', [AuthController::class, 'login']);

Route::middleware('auth:sanctum')->group(function () {
    Route::get('/user', function (Request $request) {
        return $request->user();
    });

    Route::post('/logout', [AuthController::class, 'logout']);
    Route::post('/upload', [OtdrController::class, 'upload']);
    Route::get('/history', [OtdrController::class, 'history']);
});
