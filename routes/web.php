<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExportController;
use App\Http\Controllers\ReportController;

Route::get('/', function () {
    return view('welcome');
});

Route::get('export-data', [ExportController::class, "export_data"]);
Route::get('report-export', [ReportController::class, 'index']);