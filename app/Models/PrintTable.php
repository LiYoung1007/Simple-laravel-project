<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class PrintTable extends Model
{
    use HasFactory;
    protected $table = 'tbl_points';
    protected $fillable = [
        'point_id', 'hyosyo_id', 'hyosyo_from', 'hyosyo_to' ,'no','person_id', 'point', 'messageCount', 'receivePoint','sendPoint' , 'created' , 'modified'
    ];
}
