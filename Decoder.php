<?php

require 'vendor/autoload.php';

class Decoder
{
    /**
     * Возвращает строку в массив
     * @param string $jsonEncodedString
     * @return mixed
     * @throws Exception
     */
    public static function unserializeArray(string $jsonEncodedString): array
    {
        return json_decode(gzuncompress(base64_decode(urldecode($jsonEncodedString))), true);
    }
}





