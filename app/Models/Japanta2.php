<?php
 
namespace App\Models;
 
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
 
class Japanta extends Model
{
    protected $table = 'japanta';
    use HasFactory;
 
    protected $fillable = [
        'fecha',
        'nombre',
        'grupo',
        'saldoinicial',
        'debe',
        'haber',
        'saldofinal'
    ];
}