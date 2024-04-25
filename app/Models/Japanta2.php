<?php
 
namespace App\Models;
 
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
 
class Japanta extends Model
{
    protected $table = 'japanta2';
    use HasFactory;
 
    protected $fillable = [
        'fecha2',
        'nombre',
        'grupo',
        'saldoinicial',
        'debe2',
        'haber2',
        'saldofinal'
    ];
}