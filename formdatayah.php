<?php
   include 'koneksi.inc.php';
   $ayah = mysqli_query($koneksi, "SELECT * from dataayah ORDER BY namaayah") or die("Proses registrasi gagal" );
?>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>FORMULIR PESERTA DIDIK</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.0-beta1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-0evHe/X+R7YkIZDRvuzKMRqM+OrBnVFBL6DOitfPri4tjfHxaWutUpFmBp4vmVor" crossorigin="anonymous">
    <style>
        .warning {color: #FF0000;}
    </style>
</head>
<body>

<?php
    $error_namaayah = "";
    $error_tlayah = "";
    $error_pendayah = "";
    $error_pekerayah = "";
    $error_salaryayah = "";
    $error_berkebayah = "";

    $namaayah = "";
    $tlayah = "";
    $pendayah = "";
    $pekerayah = "";
    $salaryayah = "";
    $berkebayah = "";

    if ($_SERVER["REQUEST_METHOD"] == "POST") {

        if (empty($_POST["tlayah"])) {
            $error_tlayah = "Tanggal Lahir Tidak Boleh Kosong";
        } else {
            $tlayah = cek_input($_POST["tlayah"]);
            if (!is_numeric($tlayah)) {
                $error_tlayah = "Tanggal Lahir Hanya Boleh Angka";
            }
        }

    }

    function cek_input($data) {
        $data = trim($data);
        $data = stripcslashes($data);
        $data = htmlspecialchars($data);
        return $data;
    }
    ?>
    
<div class="row">
        <div class="col-md-8">
            <div class="card">
                <div class="card-header">
                FORMULIR PESERTA DIDIK
                </div>
                <div class="card-footer">
                    3. FORM DATA AYAH KANDUNG
                </div>
                <div class="card-body">
                    <form method="post" action="prosesdatayah.php">

                    <div class="form-group row">
                        <label for="namaayah" class="col-sm-2 col-form-label">Nama Ayah Kandung</label>
                        <div class="col-sm-10">
                            <input type="text" name="namaayah" id="namaayah" class="form-control <?php echo ($error_namaayah !="" ? "is-invalid" : ""); ?>" placeholder="" value="<?php echo $namaayah; ?>"> <span class="warning"><?php echo $error_namaayah ?></span>
                        </div>
                    </div><br>

                    <div class="form-group row">
                        <label for="tlayah" class="col-sm-2 col-form-label">Tahun Lahir</label>
                        <div class="col-sm-10">
                            <input type="text" name="tlayah" id="tlayah" class="form-control <?php echo ($error_tlayah !="" ? "is-invalid" : ""); ?>" placeholder="" value="<?php echo $tlayah; ?>"> <span class="warning"><?php echo $error_tlayah ?></span>
                        </div>
                    </div><br>

                    <div class="form-group row">
                        <label for="pendayah" class="col-sm-2 col-form-label "> Pendidikan </label>
                        <div class="col-sm-10">
                            <select class="col-sm-2 form-select" name="pendayah">
                            <option value=""></option>
                            <option value="Tidak Sekolah"> Tidak Sekolah </option>
                            <option value="Putus SD"> Putus SD </option>
                            <option value="SD Sederajat"> SD Sederajat </option>
                            <option value="SMP Sederajat"> SMP Sederajat </option>
                            <option value="SMA Sederajat"> SMA Sederajat </option>
                            <option value="D1"> D1 </option>
                            <option value="D2"> D2 </option>
                            <option value="D3"> D3 </option>
                            <option value="D4/S1"> D4/S1 </option>
                            <option value="S2"> S2 </option>
                            <option value="S3"> S3 </option>
                            </select>
                            <span class="warning" ><?php echo $error_pendayah; ?></span>
                        </div>
                    </div><br>

                    <div class="form-group row">
                        <label for="pekerayah" class="col-sm-2 col-form-label "> Pekerjaan </label>
                        <div class="col-sm-10">
                            <select class="col-sm-2 form-select" name="pekerayah">
                            <option value=""></option>
                            <option value="PNS"> Tidak Bekerja </option>
                            <option value="TNI/POLRI"> Nelayan </option>
                            <option value="Guru/Dosen"> Petani </option>
                            <option value="Dokter"> Peternak </option>
                            <option value="Politikus"> PNS/TNI/POLRI </option>
                            <option value="Wiraswasta"> Karyawan Swasta </option>
                            <option value="Lainnya"> Pedagang Kecil </option>
                            <option value="Lainnya"> Pedagang Besar </option>
                            <option value="Lainnya"> Wiraswasta </option>
                            <option value="Lainnya"> Wirausaha </option>
                            <option value="Lainnya"> Buruh </option>
                            <option value="Lainnya"> Pensiunan </option>
                            <option value="Lainnya"> Lainnya </option>
                            </select>
                            <span class="warning" ><?php echo $error_pekerayah; ?></span>
                        </div>
                    </div><br>

                    <div class="form-group row">
                        <label for="salaryayah" class="col-sm-2 col-form-label">Penghasilan Bulanan</label>
                        <div class="col-sm-10">
                            <input type="radio" name="salaryayah" value="< 500.000"> < 500.000 </label> <br>
                            <input type="radio" name="salaryayah" value="500.000 - 999.9999"> 500.000 - 999.9999 </label><br>
                            <input type="radio" name="salaryayah" value="1.000.000 - 1.999.999"> 1.000.000 - 1.999.999 </label><br>
                            <input type="radio" name="salaryayah" value="2.000.000 - 4.999.999"> 2.000.000 - 4.999.999 </label><br>
                            <input type="radio" name="salaryayah" value="5.0000.000 - 20.000.000"> 5.0000.000 - 20.000.000 </label> <br>
                            <input type="radio" name="salaryayah" value="20.000.000"> > 20.000.000 </label>
                            <span class="warning" ><?php echo $error_salaryayah; ?></span>
                        </div>
                    </div><br>

                    <div class="form-group row">
                        <label for="berkebayah" class="col-sm-2 col-form-label "> Berkebutuhan Khusus </label>
                        <div class="col-sm-10">
                            <select class="col-sm-2 form-select" name="berkebayah">
                            <option value=""></option>
                            <option value="Tidak"> Tidak </option>
                            <option value="Netra (A)"> Netra (A) </option>
                            <option value="Rungu (B)"> Rungu (B) </option>
                            <option value="Grahita Ringan (C)"> Grahita Ringan (C) </option>
                            <option value="Grahita Sedang (C1)"> Grahita Sedang (C1) </option>
                            <option value="Daksa Ringan (D)"> Daksa Ringan (D) </option>
                            <option value="Laras (E)"> Laras (E) </option>
                            <option value="Wicara (F)"> Wicara (F) </option>
                            <option value="Tuna Ganda (G)"> Tuna Ganda (G) </option>
                            <option value="Hiper Aktif (H)"> Hiper Aktif (H) </option>
                            <option value="Cerdas Istimewa (I)"> Cerdas Istimewa (I) </option>
                            <option value="Bakat Istimewa (J)"> Bakat Istimewa (J) </option>
                            <option value="Kesulitan Belajar (K)"> Kesulitan Belajar (K) </option>
                            <option value="Narkoba (N)"> Narkoba (N) </option>
                            <option value="Indigo (O)"> Indigo (O) </option>
                            <option value="Down Syndrom (P)"> Down Syndrom (P) </option>
                            <option value="Autis (Q)"> Autis (Q) </option>
                            <option value="Lainnya"> Lainnya </option>
                            </select>
                            <span class="warning" ><?php echo $error_berkebayah; ?></span>
                        </div>
                    </div>

                    <br>
                    <div class="form-group row">
                        <div class="col-sm-10">
                            <button type="submit" name="submit" class="btn btn-primary">Next</button>
                        </div>
                    </div>
                    <a class="btn btn-danger" href="reportdataexcel3.php">Cetak Excel</a>
                    </form>
                </div>
            </div>
        </div>
    </div>
</body>
</html>