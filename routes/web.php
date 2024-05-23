<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ExportController;



Route::get('/', function () {
    return view('welcome');
});

Route::get('export-data', [ExportController::class, "export_data"]);