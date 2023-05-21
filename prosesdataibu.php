<?php 
include 'koneksi.inc.php';
    if (isset($_POST['submit'])) {
        $namaibu = $_POST['namaibu'];
        $tlibu = $_POST['tlibu'];
        $pendibu = $_POST['pendibu'];
        $pekeribu = $_POST['pekeribu'];
        $salaryibu = $_POST['salaryibu'];
        $berkebibu = $_POST['berkebibu'];

      // Menyimpan ke database
      $sql = mysqli_query($koneksi, "INSERT INTO dataibu (namaibu, tlibu, pendibu, pekeribu, salaryibu, berkebibu) VALUES ('$namaibu', '$tlibu', '$pendibu', '$pekeribu', '$salaryibu', '$berkebibu')");
      if ($sql) {
          // pesan jika data tersimpan
          echo "<script>alert('Data Keseluruhan Berhasil Disimpan!'); window.location.href='formregist.php'</script>"; 
        } else {
          // pesan jika data gagal disimpan
            echo "alert('Data Ibu Kandung Gagal Ditambahkan!!');";
        }
      }
?>