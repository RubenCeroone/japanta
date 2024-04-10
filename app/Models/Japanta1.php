<?php
 
namespace App\Models;
 
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
 
class Japanta1 extends Model
{
    protected $table = 'japanta1';
    use HasFactory;
 
    protected $fillable = [
        'fecha1',
        'concepto1',
        'documento1',
        'tags1',
        'debe1',
        'haber1',
        'saldo1'
    ];
}