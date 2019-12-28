Public Class Class1
    'запрос из двух таблиц

    'Dim ds = (From x In dtShtatnoeOtdelyAll.AsEnumerable
    'Join y In dtClientAll.AsEnumerable On x.Field(Of String)("Клиент") Equals
    '   y.Field(Of String)("НазвОрг")
    'Where y.Field(Of String)("НазвОрг") = ComboBox1.Text
    'Order By x.Field(Of String)("Отделы")
    'Select Case x.Field(Of String)("Отделы") Distinct)

    'добавляет недостающие нули в чмсло в строку
    'Dim f2 As Double = 38.5
    'Dim f1 As String = Format(f2, "f") 'форматирует 

    'Dim ds2 = (From x In dtDogovorPadriadaAll.AsEnumerable
    '           Join y In dtDogPodrAktInoeAll.AsEnumerable On x.Field(Of Integer)("Код") Equals
    '              y.Field(Of Integer)("IDДогПодряда")
    '           Where Not x.IsNull("ID") AndAlso x.Item("ID") = j _
    '              AndAlso Not x.IsNull("НомерДогПодр") AndAlso x.Item("НомерДогПодр") = ComboBox3.Text _
    '              AndAlso Not y.IsNull("ПорНомерАктаИное") AndAlso y.Item("ПорНомерАктаИное") = ComboBox5.Text
    '           Select New With {.Наименование = y.Item("ВыпРаб1"), .Единица = y.Item("ЕдИзмерАктИное"),
    '              .Стоимость2 = y.Item("СтоимЕдРаботыАктИное"), .Объем = y.Item("ОбъемВыпРаботАктИное"),
    '              .Стоимость = y.Item("ОбщСтоимРаботАктИное"), .Код = y.Item("ID")}).ToList()




    'LINQ

    'Private void button1_Click(Object sender, EventArgs e)  //insert
    '    {
    '        DataClasses1DataContext db = New DataClasses1DataContext();
    '        tblKisiler yenikisi = New tblKisiler();
    '        yenikisi.ad = textBox2.Text;
    '        yenikisi.soyad = textBox3.Text;
    '        yenikisi.telefon = textBox4.Text;
    '        db.tblKisilers.InsertOnSubmit(yenikisi);
    '        db.SubmitChanges();
    '        dataGridView1.DataSource = db.tblKisilers;
    '    }

    '    Private void button2_Click(Object sender, EventArgs e)  //update
    '    {
    '        db = New DataClasses1DataContext();
    '        int sayi = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
    '        var guncelle = db.tblKisilers.Where(w =& gt; w.id == sayi).FirstOrDefault();
    '        guncelle.ad = textBox2.Text;
    '        guncelle.soyad = textBox3.Text;
    '        guncelle.telefon = textBox4.Text;
    '        db.SubmitChanges();
    '        dataGridView1.DataSource = db.tblKisilers;
    '    }
    'Private void button3_Click(Object sender, EventArgs e) //delete
    '    {
    '        db = New DataClasses1DataContext();
    '        int sayi = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
    '        var sil = db.tblKisilers.Where(w =& gt; w.id == sayi).FirstOrDefault();
    '        db.tblKisilers.DeleteOnSubmit(sil);
    '        db.SubmitChanges();
    '        dataGridView1.DataSource = db.tblKisilers;

    '    }


    'Using dbcx = New DbAllDataContext  'мой insert
    'Dim f As New ДогПодрОбязан()
    '        f.Обязанности = ComboBox1.Text
    '        f.ID = RichTextBox2.Text
    '        dbcx.ДогПодрОбязан.InsertOnSubmit(f)
    '        dbcx.SubmitChanges()
    '        idДолжность = f.Код
    '    End Using



    'Using dbcx = New DbAllDataContext() 'мой update
    'Dim var = (From x In dbcx.ДогПодрОбязан.AsEnumerable Where x.Код = idОбязанность Select x).Single
    'If var IsNot Nothing Then
    '            var.Обязанности = RichTextBox1.Text
    '            dbcx.SubmitChanges()
    '        End If
    'End Using


    'Using dbcx = New DbAllDataContext() 'мой update2
    'Dim var = dbcx.ДогПодДолжн.Single(Function(x) x.Код = idДолжность)
    'If var IsNot Nothing Then
    '            var.Должность = RichTextBox2.Text
    '            dbcx.SubmitChanges()
    '        End If
    'End Using


    'Using dbcx = New DbAllDataContext() 'мой delete
    'Dim var = dbcx.ДогПодДолжн.Single(Function(x) x.Код = ComboBox22.SelectedValue)
    'If var IsNot Nothing Then
    '            dbcx.ДогПодДолжн.DeleteOnSubmit(var)
    '            dbcx.SubmitChanges()

    '        End If
    'End Using








    'Dim ds1   'комбобокс с id
    '    dbcx = New DbAllDataContext
    '    ds1 = From x In dbcx.ДогПодДолжн Where x.Клиент = ComboBox1.Text
    '          Order By x.Должность
    '          Select x.Должность, x.Код

    '    ComboBox22.DataSource = ds1
    '    ComboBox22.DisplayMember = "Должность"
    '    ComboBox22.ValueMember = "Код"







End Class
