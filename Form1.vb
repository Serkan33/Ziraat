Imports System.Data.SqlClient
Imports System.Data
Imports System.Threading
Public Class Form1
    Dim baglanti As New SqlConnection("Data Source=(LocalDB)\v11.0;Integrated Security=True;AttachDbFileName='|DataDirectory|\ziraatOto.mdf'")
    Dim baglanti2 As New SqlConnection()
    Dim komut As New SqlCommand
    Dim ekleSorgu As String
    Dim dogumTarihi As String
    Dim adapter As New SqlDataAdapter("SELECT * FROM Kisiler where tcno is not null", baglanti)
    Dim dataTable As New DataTable
    Dim dataSet As New DataSet
    Dim id As Integer
    Dim gelenId As Integer
    Dim zSebep(24) As CheckBox

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Burada dogum comboboxlarını dolduran döngüleri çalıştırıyorum
        For index = 1 To 31
            combogun.Items.Add(index)
        Next

        For index = 1950 To 2016
            comboyil.Items.Add(index)
        Next
        'zSebep CheckBox tipinde bir dizi ve kullanacağım checkboxları bu dizide tutuyorum
        zSebep(0) = CheckBox1
        zSebep(1) = CheckBox2
        zSebep(2) = CheckBox3
        'goster fonksiyonu veritabanından tüm verileri çekip gösteriyor
        goster()


    End Sub
    Public Sub goster()
        ' gösterme işlemleri burada yapılıyor. Her güncelemede tabloyu silip verileri tekrardan tabloya ekliyorum
        dataTable.Clear()
        adapter.Fill(dataTable)
        DataGridView1.DataSource = dataTable
        ComboBox5.Text = "Tc Kimlik" 'arama kısmı için default değer olarak tc seçiyorum
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles eklebuton.Click

        dogumTarihi = combogun.Text + " " + comboay.Text + " " + comboyil.Text
        Try
            If textad.Text IsNot "" And textsoyad.Text IsNot "" And textbabaad.Text IsNot "" And dogumTarihi IsNot "" And comboikamet.Text IsNot "" And datekayit.Text IsNot "" And textkayit.Text IsNot "" And textklasor.Text IsNot "" And RichTextBox1.Text IsNot "" And tcText.Text IsNot "" Then
                If tcText.TextLength = 11 Then


                    Dim kontrol As String = "SELECT * FROM Kisiler WHERE tcno=@tc"
                    baglanti2 = baglanti
                    baglanti2.Open()
                    Dim cmd1 As SqlCommand = New SqlCommand(kontrol, baglanti2)
                    cmd1.Parameters.AddWithValue("@tc", tcText.Text)
                    'kontrol değişenine veritanında tc soruluyorum.Eğer aynı tc varsa uyarı veriyor. Eğer yoksa veriyi ekliyor
                    Using reader As SqlDataReader = cmd1.ExecuteReader()
                        If reader.HasRows = False Then
                            baglanti2.Close()
                            komut.Connection = baglanti
                            'Aşağıdaki kodda aldığım verleri veritabanına ekliyorum
                            komut.CommandText = "insert into Kisiler (kisi_ad,kisi_soyad,kisi_baba,kisi_dogumtarihi,ikamet,kayit_tarihi,kayit_no,klasor,aciklama,tcno)" +
                        "values('" & textad.Text & "','" & textsoyad.Text & "','" & textbabaad.Text & "','" & dogumTarihi & "','" & comboikamet.Text & "','" & datekayit.Text & "','" & textkayit.Text & "','" & textklasor.Text & "','" & RichTextBox1.Text & "','" & tcText.Text & "')"
                            baglanti.Open()
                            komut.ExecuteNonQuery() ' Komudu burada işliyorum
                            komut.CommandText = "select top(1) kisi_id from Kisiler order by kisi_id desc " ' burada ise en son eklenen kişinin id sini alıyorum. Bunu güncellemede kullanacağım için yapıyorum
                            Using sec As SqlDataReader = komut.ExecuteReader
                                If sec.HasRows Then
                                    Do While sec.Read()
                                        gelenId = sec.GetInt32(0) ' Çektiğim id yi gelenId değişkenine atıyorum

                                    Loop
                                End If
                            End Using
                            '----------------------------------------------------------------------------------------------------------------
                            If CheckBox1.Checked Then
                                'Eğer CKS işaretli ise veritabanına ekliyorum 
                                Try
                                    komut.CommandText = "insert into ziyaret(CKS,kisi_id) values('" & CheckBox1.Text & "','" & gelenId & "')"
                                    komut.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("Ziyarette Hata Var")
                                End Try

                            End If
                            If CheckBox3.Checked Then
                                Try
                                    'Eğer CKS işaretli ise veritabanına ekliyorum
                                    komut.CommandText = "insert into ziyaret(FKS,kisi_id) values('" & CheckBox3.Text & "','" & gelenId & "')"
                                    komut.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("Ziyarette Hata Var")
                                End Try

                            End If
                            If CheckBox2.Checked Then
                                Try
                                    'Eğer MGD işaretli ise veritabanına ekliyorum
                                    komut.CommandText = "insert into ziyaret(MGD,kisi_id) values('" & CheckBox2.Text & "','" & gelenId & "')"
                                    komut.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("Ziyarette Hata Var")
                                End Try

                            End If
                            If CheckBox4.Checked Then
                                Try
                                    'Eğer OKS işaretli ise veritabanına ekliyorum
                                    komut.CommandText = "insert into ziyaret(OKS,kisi_id) values('" & CheckBox4.Text & "','" & gelenId & "')"
                                    komut.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show("Ziyarette Hata Var")
                                End Try

                            End If
                            '----------------------------------------------------------------------------------------------------------------
                            'Ekleme işlemeleri başarı olmuşsa aşağıdaki uyarıyı veriyorum
                            MessageBox.Show("Üye Başarılı Bir Şekilde Eklendi", "İŞLEM BAŞARILI", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                            goster() ' goster fonksiyonunu çalıştırarak tabloda ki verileri güncel tutmaya çalışıyorum
                            baglanti.Close()
                            'Aşağıdaki kodalr ise formdaki bütün alanları temizliyor
                            tcText.Clear()
                            textad.Clear()
                            textbabaad.Clear()
                            textkayit.Clear()
                            textklasor.Clear()
                            textsoyad.Clear()
                            comboay.SelectedItem = ""
                            combogun.SelectedItem = ""
                            comboyil.SelectedItem = ""
                            comboikamet.SelectedItem = ""
                            RichTextBox1.Clear()
                            CheckBox1.Checked = False
                            CheckBox2.Checked = False
                            CheckBox3.Checked = False
                            CheckBox4.Checked = False

                        Else
                            'Kontrol ettiğimiz tc veritabanında varsa o kişiyi eklemeyip uyarı veriyorum
                            MsgBox("Üye Sistemde Zaten Mevcut", MsgBoxStyle.Exclamation, "Yeni Üye Ekle")
                            baglanti2.Close()

                        End If



                    End Using
                ElseIf tcText.TextLength > 11 Or tcText.TextLength < 11 Then
                    'tc alanını 11 karakterli olması gerektiğini kontrol ediyorum
                    MsgBox("TC Kimlik 11 rakamdan oluşmalıdır", MsgBoxStyle.Critical, "Hata")
                Else
                    'boş alanlara izin vermiyorum
                    MsgBox("Lütfen Bütün Alanları Doldurunuz", MsgBoxStyle.Critical, "Hata")
                End If
            End If

        Catch ex As Exception
            'Ekleme işlemlerinde hata olursa aşağıdaki uyarıyı veriyorum
            MessageBox.Show("Hata")
        End Try




    End Sub

    Private Sub duzenlebuton_Click(sender As Object, e As EventArgs) Handles duzenlebuton.Click

        'Bu alanda güncelleme(düzenleme işlemlerini yapıyorum)

        'dogum tarihini alıyorum
        If combogun.Text IsNot "" And comboay.Text IsNot "" And comboyil.Text IsNot "" Then
            dogumTarihi = combogun.Text + " " + comboay.Text + " " + comboyil.Text
        End If

        Try
            'Kullanıcının bigilerini alıp handi kullanıcıya aitse o bilgileri yeni bilgilerle değiştiriyorum
            komut.Connection = baglanti
            komut.CommandText = "update kisiler set kisi_ad='" & textad.Text & "',kisi_soyad='" & textsoyad.Text & "',kisi_baba='" & textbabaad.Text & "',ikamet='" & comboikamet.Text & "',kayit_tarihi='" & datekayit.Text & "',kayit_no='" & textkayit.Text & "',klasor='" & textklasor.Text & "',aciklama='" & RichTextBox1.Text & "',tcno='" & tcText.Text & "' where kisi_id='" & id & "' "
            baglanti.Open()
            komut.ExecuteNonQuery()
            'Dogum tarihi alanı boş değilse dogum tarihini ekliyorum
            If dogumTarihi.Trim() IsNot "" Then

                komut.CommandText = "update kisiler set kisi_dogumtarihi='" & dogumTarihi & "' where kisi_id='" & id & "' "
                komut.ExecuteNonQuery()
            End If


            If CheckBox1.Checked = False Then
                'Aşagıdaki işlemlerde işaretlenmemiş veriler varsa onları veritabanında bos olarak güncelliyorum.Eğer null olarak dönerse hata oluşur
                Try
                    komut.CommandText = "update ziyaret set CKS='bos' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t1")
                End Try

            End If
            If CheckBox2.Checked = False Then
                Try
                    komut.CommandText = "update ziyaret set MGD='bos' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t2")
                End Try

            End If
            If CheckBox3.Checked = False Then
                Try
                    komut.CommandText = "update ziyaret set FKS='bos' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t3")
                End Try

            End If
            If CheckBox4.Checked = False Then
                Try
                    komut.CommandText = "update ziyaret set OKS='bos' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t4")
                End Try

            End If
            '-------------
            'Burada eğer işaretlenmiş kısımlar varsa onları kendi isimlerinde veritabanında güncelliyorum
            If CheckBox1.Checked Then
                Try
                    komut.CommandText = "update ziyaret set CKS='" & CheckBox1.Text & "' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t5")
                End Try

            End If
            If CheckBox3.Checked Then
                Try
                    komut.CommandText = "update ziyaret set FKS='" & CheckBox3.Text & "' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t6")
                End Try

            End If
            If CheckBox2.Checked Then
                Try
                    komut.CommandText = "update ziyaret set MGD='" & CheckBox2.Text & "' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t7")
                End Try

            End If
            If CheckBox4.Checked Then
                Try
                    komut.CommandText = "update ziyaret set OKS='" & CheckBox4.Text & "' where kisi_id='" & id & "' "
                    komut.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("t8")
                End Try

            End If
            goster()
            MessageBox.Show("Başarılı Bir Şekilde Güncellendi", "İŞLEM BAŞARILI", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

        Catch ex As Exception
            MessageBox.Show("Doğum Tarihi Alanını Kontrol Ediniz")
        End Try

        baglanti.Close()


    End Sub

    Private Sub silbuton_Click(sender As Object, e As EventArgs) Handles silbuton.Click

        'Bu alanda veri silme işlemi gerçekleşiyor. Kullanıcının id sine göre silme komudunu uyguluyorum
        If MessageBox.Show("Seçtiğiniz Bilgiyi Kalıcı Olarak Silmek İstediğinizden Emin misiniz?", "UYARI", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) = DialogResult.Yes Then
            Try
                komut.Connection = baglanti
                komut.CommandText = "delete from kisiler where kisi_id='" & id & "' "
                komut.CommandText = "delete from kisiler where kisi_id='" & id & "' "
                baglanti.Open()
                komut.ExecuteNonQuery()
                goster()

                MessageBox.Show("Silme İşlemi Başarılı", "İŞLEM BAŞARILI", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)

            Catch ex As Exception

            End Try

        End If



        baglanti.Close()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        'Burada ise taboda hangi satıra tıklarsak o satır ait bütün verileri alıp ilgili alanlara aktarıyorum. Aşağıda ki işlemler bunu yapıyor. Ve kullanıcının id sini bu şekilde alıyoruz.
        Dim satir As Integer
        'Butona tıklanınca tablomuzun geçerli satırının indexini(değerini) alıyorum
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
        satir = DataGridView1.CurrentCell.RowIndex
        tcText.Text = DataGridView1(1, satir).Value
        textad.Text = DataGridView1(2, satir).Value
        textsoyad.Text = DataGridView1(3, satir).Value
        textbabaad.Text = DataGridView1(4, satir).Value
        comboikamet.Text = DataGridView1(6, satir).Value
        datekayit.Text = DataGridView1(7, satir).Value
        textkayit.Text = DataGridView1(8, satir).Value
        textklasor.Text = DataGridView1(9, satir).Value
        RichTextBox1.Text = DataGridView1(10, satir).Value
        id = DataGridView1(0, satir).Value
        komut.Connection = baglanti
        komut.CommandText = "select CKS,MGD,OKS,FKS  from ziyaret where kisi_id='" & id & "'"
        baglanti.Open()
        Using sec As SqlDataReader = komut.ExecuteReader
            If sec.HasRows Then
                Do While sec.Read()
                    If CheckBox1.Text = sec.GetString(0) Then
                        CheckBox1.Checked = True
                    End If

                    If CheckBox2.Text = sec.GetString(1) Then
                        CheckBox2.Checked = True
                    End If
                    If CheckBox4.Text = sec.GetString(2) Then
                        CheckBox4.Checked = True
                    End If
                    If CheckBox3.Text = sec.GetString(3) Then
                        CheckBox3.Checked = True
                    End If

                Loop
            End If
        End Using
        baglanti.Close()
    End Sub

    Private Sub tcText_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tcText.KeyPress
        If Not (Char.IsNumber(e.KeyChar) = True) And e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Dim filtre As String
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged

        ' Arama işlemini burada geçekleştiriyorum
        ' Aramaya ait combobox değeri değiştiğinde o değeri alıp filtre değişkenine atıyorum. 
        Dim dv As New DataView(dataTable)
        If filtre = "Tc Kimlik" Then
            filtre = "tcno"
        ElseIf filtre = "Ad" Then
            filtre = "kisi_ad"
        ElseIf filtre = "Soyad" Then
            filtre = "kisi_soyad"
        ElseIf filtre = "Kayıt No" Then
            filtre = "kayit_no"
        ElseIf filtre = "Klasor" Then
            filtre = "klasor"
        End If

        'aldığım filtre değişkenine göre tabloda aramayı yaptırıyorum
        dv.RowFilter = String.Format(" {0} LIKE '%" & TextBox6.Text & "%'", filtre)
        DataGridView1.DataSource = dv

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        'Aramaya ait combobox burada ve bunun değeri değişince o değeri filtre değişkeninde tutuyorum
        filtre = ComboBox5.Text
        TextBox6.Text = ""
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_3(sender As Object, e As EventArgs)
        MessageBox.Show(String.Format(zSebep(0).Text))
        String.Format(zSebep(0).Text)
    End Sub

    Private Sub Button1_Click_4(sender As Object, e As EventArgs) Handles Button1.Click
        'Formdaki Bütün alanları temizliyorum
        tcText.Clear()
        textad.Clear()
        textbabaad.Clear()
        textkayit.Clear()
        textklasor.Clear()
        textsoyad.Clear()
        comboay.SelectedItem = ""
        combogun.SelectedItem = ""
        comboyil.SelectedItem = ""
        comboikamet.SelectedItem = ""
        RichTextBox1.Clear()
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        CheckBox4.Checked = False
    End Sub
End Class