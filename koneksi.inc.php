<?php
$host = "localhost";
$user = "root";
$password = "";
$database = "latihan";

// Membuat Koneksi
$koneksi = mysqli_connect($host, $user, $password, $database) or die("Koneksi gagal dibangun dan Database tidak dapat dibuka");
?>