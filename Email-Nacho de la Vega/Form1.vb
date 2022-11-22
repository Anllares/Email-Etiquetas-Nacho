Imports System.Diagnostics
Imports System.Reflection 
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Text
Imports System.Net.Mail
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Net.Mime


Public Class Form1
	Dim oDocument As Object	
	Dim sFichero As String '' fichero word o html para generar el mailing
	Dim condicionfiltro As String	
	Dim matrizprovincias() As DataRow
	Dim pararproceso As Boolean
	Dim haycontacto As Boolean
	
	Sub ButtonSepararDatosClick(sender As Object, e As EventArgs)
		
		Me.Cursor=System.Windows.Forms.Cursors.WaitCursor
		Try
			DataGridView1.Columns.Clear()
			
            For Each Col As DataGridViewColumn In DataGridViewDatos.Columns
            	DataGridView1.Columns.Add(DirectCast(Col.Clone, DataGridViewColumn))            
            Next
            For rowIndex As Integer = 0 To (DataGridViewDatos.Rows.Count - 1)
            	'MessageBox.Show(DataGridViewDatos.Rows(rowIndex).Cells("CPOBCLI").Value.ToString)
            	'' If instr(1,condicionfiltro,DataGridViewDatos.Rows(rowIndex).Cells("CCODGRUP").Value.ToString,1) Then           	
            	 	If instr(1,condicionfiltro,DataGridViewDatos.Rows(rowIndex).Cells("CCODIGO").Value.ToString,1) Then           	
                    	DataGridView1.Rows.Add(DataGridViewDatos.Rows(rowIndex).Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray)
                    End if	
				
            Next
            
		Catch ex As Exception
		End Try		
		Me.textBoxFiltrados.Text="Registros: " & DataGridView1.RowCount.ToString
		Me.Cursor=System.Windows.Forms.Cursors.Default
End Sub


Public Shared Function encode(ByVal str As String) As String
    'supply True as the construction parameter to indicate
    'that you wanted the class to emit BOM (Byte Order Mark)
    'NOTE: this BOM value is the indicator of a UTF-8 string
    
   ' Dim textOEncoding As New System.Text.UTF8Encoding()
    
    Dim textoEncoding As New System.Text.ASCIIEncoding()
    Dim encodedString() As Byte

    encodedString = textoEncoding.GetBytes(str)

    Return textoEncoding.GetString(encodedString)
End Function


Public Shared Function Transforma(byval str As String) As String
    Dim encodestring As String
    
    If IsNothing(str) Then    	
    	Return ""
    End If
    
    If str.Length<=0 Then
    	Return ""
    End If
    
    
    'AA EE II OO UU ÑÑ ÞÜ ºª ÇÇ
    
    encodestring=Trim(str)
    
    encodestring=Replace(encodestring,"Þ","Ü") 
    encodestring=Replace(encodestring,"Ã","Ü")
        
    encodestring=Replace(encodestring,"┴","A")
    encodestring=Replace(encodestring,"ß","A")
    encodestring=Replace(encodestring,"╔","E")
    encodestring=Replace(encodestring,"Ú","E")
    encodestring=Replace(encodestring,"═","I")
    encodestring=Replace(encodestring,"Ý","I")
    encodestring=Replace(encodestring,"Ë","O")
    encodestring=Replace(encodestring,"¾","O")
    encodestring=Replace(encodestring,"┌","U")
    encodestring=Replace(encodestring,"·","U")
    
    encodestring=Replace(encodestring,"Ð","Ñ")
    encodestring=Replace(encodestring,"±","Ñ")
    
    encodestring=Replace(encodestring,"║","º")
    encodestring=Replace(encodestring,"¬","ª")
    
        encodestring=Replace(encodestring,"▄","Ç")
    encodestring=Replace(encodestring,"³","Ç")
    
    Return encodeString
End Function

Function ValidateEmail(ByVal email As String) As Boolean
    Dim emailRegex As New System.Text.RegularExpressions.Regex( 
        "^(?<user>[^@]+)@(?<host>.+)$")
    Dim emailMatch As System.Text.RegularExpressions.Match = 
       emailRegex.Match(email)
    Return emailMatch.Success
End Function

    Sub ButtonEnviarClick(sender As Object, e As EventArgs)
    	    	
    	Dim auxprov,nprovincia As String
    	Dim contaenvios As Integer=0    	    	
    	Dim nombre,contacto As String
		Dim direccion As String
		Dim poblacion As String
		Dim cpostal As String
		Dim email As String	
		Dim seleccionada As String
		Dim esrojo As Boolean
		
		Dim oWord As Microsoft.Office.Interop.Word.Application ''Word.Application
    	Dim oDoc As Microsoft.Office.Interop.Word.Document ''Word.Document

    	
    	If Me.checkBoxCarta.Checked=False And Me.checkBoxEmail.Checked=False Then			
    		MessageBox.Show("No has elegido qué enviar, carta o e-mail","Error...",MessageBoxButtons.OK,MessageBoxIcon.Warning)
    		return
        End If
        
        
        Me.AxwebBrowser1.Visible=False
        
        '' añado las columnas
        Me.dataGridViewEnvios.Columns.Clear()
        Me.dataGridViewEnvios.ColumnCount=5
        Me.dataGridViewEnvios.Columns(0).Name="Número"
        Me.dataGridViewEnvios.Columns(1).Name="Nombre"
        Me.dataGridViewEnvios.Columns(2).Name="Dirección"
        Me.dataGridViewEnvios.Columns(3).Name="Email"
        Me.dataGridViewEnvios.Columns(4).Name="Código Postal"
        Me.dataGridViewEnvios.Location=New Point (6,177)
        Me.dataGridViewEnvios.Visible=True
        Me.dataGridViewEnvios.Height=Me.AxwebBrowser1.Height
        Me.dataGridViewEnvios.Width=Me.AxwebBrowser1.Width
        Me.dataGridViewEnvios.AutoSizeColumnsMode=DataGridViewAutoSizeColumnMode.Fill
        
'        Me.textBoxEnvio.Location=New Point (6,177)
        'Me.textBoxEnvio.AutoSize=false
'        Me.textBoxEnvio.Height=Me.AxwebBrowser1.Height
'        Me.textBoxEnvio.Width=Me.AxwebBrowser1.Width
       ' Me.textBoxEnvio.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        
		'Me.textBoxEnvio.Visible=True        
		
		
		Me.buttonSalir.Enabled=False
		'Me.DataGridView1.Enabled=False
		
   		      
    	'desactivo el botón de enviar
    	Me.ButtonElegirFichero.Enabled=False
    	Me.ButtonCargarDatos.Enabled=False
    	Me.ButtonSepararDatos.Enabled=False
    	Me.buttonEnviar.Enabled=False
    	Me.buttonDatosProveedores.Enabled=False
		Dim m As Object = System.Reflection.Missing.Value        
		
		Dim i As Integer
			
			
		pararproceso=False
		Me.buttonParar.Enabled=True
			
			
		Me.dataGridViewEnvios.Focus
			
			
		If Me.checkBoxCarta.Checked Then '' arranco word una vez        
			   		'Start Word and open the document template.
		 			'Dim oWord As Word.Application
    	    		'Dim oDoc As Word.Document
      	     		Try
		    	    	oWord = CreateObject("Word.Application")
			    	    oWord.Visible = false '' no se ve el word
        			Catch ex As Exception
						Throw  new Exception("Excption while trying to CreateEventSource for the EventLog: " & ex.Message)
						'MessageBox.Show(ex.Message)
					End Try
		End If			
		
		For Each row As DataGridViewRow In dataGridView1.Rows
				
				esrojo=False		
				seleccionada=row.Index
				
				'' para capturar los eventos...
				Application.DoEvents()
				If pararproceso Then					
					'' salgo del proceso
					Exit for		
				End If
				nprovincia=""
				
				If row.Cells("CPTLCLI").Value.ToString.Length=4 Then
					auxprov="0" & Strings.left(row.Cells("CPTLCLI").Value.ToString,1)
				Else
					auxprov=Strings.left(row.Cells("CPTLCLI").Value.ToString,2)					
				End If

				auxprov="00" & auxprov
			
		
    	  		''For i=0 To matrizprovincias.getupperbound(0)
    	  		''	If matrizprovincias(i)("CCODPROV").ToString=auxprov Then
    	  		''		nprovincia=matrizprovincias(i)("CNOMPROV").ToString
    	  				
    	  		''	End If      		
    	  		''Next
				nprovincia=row.Cells("Provincia").Value.ToString.ToUpper

				nombre=row.Cells("CNOMCLI").Value.ToString.ToUpper
				nombre=Transforma(nombre)
				
				direccion=row.Cells("CDIRCLI").Value.ToString.ToUpper
				direccion=Transforma(direccion)
				
				poblacion=row.Cells("CPOBCLI").Value.ToString.ToUpper
				poblacion=Transforma(poblacion)
				
				cpostal=row.Cells("CPTLCLI").Value.ToString.ToUpper
				
				If cpostal.ToUpper.Length=4 Then
					cpostal="0" & cpostal
				End If
				
				'' si hay datos de contacto
				If haycontacto Then					
					contacto=row.Cells("CCONTACTO").Value.ToString.ToUpper
					contacto=Transforma(contacto)
				Else
					contacto=""
				End if	
				
				email=row.Cells("EMAIL").Value.ToString
				
				
				If Me.checkBoxCarta.Checked Then        
			   		'Start Word and open the document template.
		 			'Dim oWord As Word.Application
    	    		'Dim oDoc As Word.Document
      	
		      		Try
      		
			    	    oWord = CreateObject("Word.Application")
			    	    oWord.Visible = false '' no se ve el word
        
        		
				
						'Me.textBoxEnvio.Text= "Cargando cargado..."        	
			   	    	oWord.Documents.Open(sfichero,m, m, m, m, m, m, m, m, m, m, m)
		    		    oDoc=oWord.ActiveDocument
    	    			'Me.textBoxEnvio.Text= "Documento cargado..."
				    	'Find and replace some text  
				        'Replace 'VB' with 'Visual Basic'  
				        oDoc.Content.Find.Execute(FindText:="[[NOMBRE]]", ReplaceWith:=nombre , Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
		        
				        oDoc.Content.Find.Execute(FindText:="[[DIRECCION]]", ReplaceWith:=direccion , Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
				        oDoc.Content.Find.Execute(FindText:="[[POBLACION]]", ReplaceWith:=poblacion, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
				        oDoc.Content.Find.Execute(FindText:="[[CPOSTAL]]", ReplaceWith:=cpostal , Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
				        oDoc.Content.Find.Execute(FindText:="[[PROVINCIA]]", ReplaceWith:=nprovincia, Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
	        
				        oDoc.Content.Find.Execute(FindText:="[[CONTACTO]]", ReplaceWith:=contacto , Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
				        oDoc.Content.Find.Execute(FindText:="[[EMAIL]]", ReplaceWith:=email , Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll)
	        
				        While oDoc.Content.Find.Execute(FindText:="  ", Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)  
				            oDoc.Content.Find.Execute(FindText:="  ", ReplaceWith:=" ", Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll, Wrap:=Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue)  
				        End While 
				        If direccion.Length>0 And poblacion.Length>0 And cpostal.Length>0 And nprovincia.length>0 Then
					    	oDoc.PrintOut()        	
		        		End If
		        		If direccion.Length<=0 Or poblacion.Length<=0 Or cpostal.Length<=0 Or nprovincia.length<=0 Then		        			
		        			esrojo=True
		        		End If
		        		
	    	    	'MessageBox.Show("Carta en Impresora","Imprimir")        	
		    	    	oDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
		    	    	oDoc=Nothing
		    	    	'Me.textBoxEnvio.Focus
		    	    	
		    	    	''oWord.Quit()
		    	    	''oWord=Nothing
    	        	    
					Catch ex As Exception
						Throw  new Exception("Excption while trying to CreateEventSource for the EventLog: " & ex.Message)
					'MessageBox.Show(ex.Message)
					End Try        
	   			End If '' de enviar carta
	   			
	   			If Me.checkBoxEmail.Checked Then
   					pararproceso=False
					Me.buttonParar.Enabled=True
										
					If email.Length ''And instr(email,"@")>0 And ValidateEmail(email) Then
						EnviarEmail(email,sFichero,nombre,textBoxSubjectEmail.Text.ToString,textBoxEnvio.text.ToString)
						''EnviarEmail("beni@nachodelavega.es",sFichero,nombre,textBoxSubjectEmail.Text.ToString,textBoxEnvio.ToString)						
					Else
						esrojo=True
					End If	        		   			
    			End If	'' de enviar email
	   			
    	    	contaenvios=contaenvios+1
    	    	'Me.textBoxEnvio.Text=contaenvios.ToString & " -> Envio a " & row.Cells("CNOMCLI").Value.ToString  & vbCrLf & Me.textBoxEnvio.Text
    	    	dataGridViewEnvios.rows.Add(New String(){contaenvios.ToString, nombre, direccion,email,cpostal})
    	    	If contaenvios Mod 100 = 0 Then
    	    		'Me.textBoxEnvio.Text="Pausa de 10 segundos..." & vbCrLf & Me.textBoxEnvio.Text    	    		    	    		
    	    		dataGridViewEnvios.rows.Add(New String(){"Pausa de 10 segundos...","","","",""})    	  
    	    		Me.DataGridViewEnvios.FirstDisplayedScrollingRowIndex = Me.DataGridViewEnvios.RowCount - 1    	    		
    	    		Me.dataGridViewEnvios.Refresh()
					System.Threading.Thread.Sleep(10000) '' pausa de un segundo    	    		
    	    	End If    	    		
    	    	
    	    	If esrojo Then
    	    	'If direccion.Length<=0 or poblacion.Length<=0 or cpostal.Length<=0 or nprovincia.length<=0 Then		        	
    	    		Me.dataGridViewEnvios.Rows(Me.dataGridViewEnvios.RowCount-2).DefaultCellStyle.ForeColor=Color.Red
    	    	Else    	    	
    	    		Me.dataGridViewEnvios.Rows(Me.dataGridViewEnvios.RowCount-2).DefaultCellStyle.ForeColor=Color.black    	    	
        		End If
				
    	    	Me.DataGridViewEnvios.FirstDisplayedScrollingRowIndex = Me.DataGridViewEnvios.RowCount - 1
    	    	
        Next
        
		If Me.checkBoxCarta.Checked Then '' arranco word una vez        			   		
    	    		Try
		    	    	''oWord.Quit()
		    	    	''oWord=Nothing
    	    			
        			Catch ex As Exception
						Throw  new Exception("Problemas creando un evento: " & ex.Message)
					End Try
		End If			
        
	   ''AxwebBrowser1.Navigate("about:blank")
	   If pararproceso Then
	   		MessageBox.Show("Proceso interrumpido por el usuario","Fin de proceso")		   	
	   Else	
			MessageBox.Show("Fin de proceso","Fin de proceso")		   	
	   End If
		
		Me.ButtonElegirFichero.Enabled=true
        Me.ButtonCargarDatos.Enabled=true
        Me.ButtonSepararDatos.Enabled=True
        Me.buttonSalir.Enabled=True
        Me.DataGridView1.Enabled=True
        Me.buttonDatosProveedores.Enabled=True
        Me.buttonEnviar.Enabled=True
        
    End Sub
    
    Sub ButtonElegirFicheroClick(sender As Object, e As EventArgs)
	Dim strFileName As String
	Dim ficherotemporal As String	
	
	
	
	Me.textBoxSubjectEmail.Visible=False
	Me.labelSubjectEmail.Visible=False
	
	
	Me.textBoxEnvio.Visible=False
	Me.dataGridViewEnvios.Visible=False
	Me.dataGridViewEnvios.Rows.Clear()
	
	Me.textBoxMensajes.Text=""
	strFileName=""
	
	Me.checkBoxCarta.Checked=False
    Me.checkBoxEmail.Enabled=False
    Me.checkBoxEmail.Checked=False
    Me.checkBoxCarta.Enabled=False    			
    Me.buttonPrueba.Enabled=False
	
	'' limpio el navegador
	AxwebBrowser1.Navigate("about:blank")
	Me.AxwebBrowser1.Visible=True
	'' //support.microsoft.com/es-es/kb/304643
    'Find the Office document.
    With OpenFileDialog1
    	.Filter = "doc files (*.doc)|*.doc|docx file (*.docx)|*.docx|html file (*.htm*)|*.htm*"
        .FileName = ""
        .ShowDialog()
        strFileName = .FileName         
        ''''sFichero=strFileName
        
    	End With
    	
    	
    	'If the user does not cancel, open the document.
    	If strFileName.Length Then
    		oDocument = Nothing	
    		'' lo paso a mayúsculas
    		strFileName=strFileName.ToUpper
    		
    		Me.Cursor=System.Windows.Forms.Cursors.WaitCursor
    		
    		ficherotemporal="\mailing\temporal\temporal" & Replace(Replace(DateTime.now.ToString,"/","_"),"-","_") 
    		ficherotemporal=Replace(ficherotemporal,":","_")
    		    		
    		If InStr(strFileName,".DOC")<>0 Then
    			ficherotemporal="C:" & Replace(ficherotemporal," ","_") & ".DOC"
    			Me.checkBoxCarta.Checked=True
    			Me.textBoxSubjectEmail.Visible=False
    			Me.labelSubjectEmail.Visible=False
    			Me.textBoxEnvio.Visible=False
			else
    			ficherotemporal="C:" & Replace(ficherotemporal," ","_") & ".HTM"
    			Me.checkBoxEmail.Checked=True    
    			Me.textBoxSubjectEmail.Visible=True
    			Me.labelSubjectEmail.Visible=True
    			Me.textBoxEnvio.Visible=False '' cuando lo arregle debe ser true
    			Me.textboxenvio.Text="PRUEBA" '' aqui se mete el texto del cuerpo del email
			End if    			
    		'MessageBox.Show(ficherotemporal,"fichero")
    		'ficherotemporal="c:\mailing\temporal\temporal.doc"
    		System.IO.File.Copy(strFileName,ficherotemporal,True)    		    		
    		'AxwebBrowser1.CausesValidation=False    		
    		'AxwebBrowser1.ScriptErrorsSuppressed=False
    		
    		
    		
    		AxWebBrowser1.Navigate(ficherotemporal,False)
    		SendKeys.send("{LEFT}")
    		SendKeys.send("{LEFT}")    		
    		'SendKeys.Send("+B")
    		'SendKeys.Send("B")
    		'SendKeys.Send("{ENTER}")
    		
    		Me.Cursor=System.Windows.Forms.Cursors.Default
    		
    		'Dim correcto As Integer=MessageBox.Show("Es el fichero a usar","Fichero",MessageBoxButtons.YesNo,MessageBoxIcon.Question)
    		
    		'If correcto=DialogResult.Yes Then
    			sFichero=strFileName
    			'AxwebBrowser1.Navigate("about:blank")
    	'	End If
    		Me.textBoxMensajes.Text="Fichero: " & sFichero    			    		    		
    		Me.buttonEnviar.Enabled=True
    		Me.buttonPrueba.Enabled=True
    	  End If    	
    End Sub
       
       Sub ButtonCargarDatosClick(sender As Object, e As EventArgs)
       	
       	'' AHORA SE CARGAR DESDE UN FICHERO EXCEL EXPORTADO DE SAGE 50 PREMIUM (LISTADOS,CLIENTES, TODOS, EN LOS DATOS QUE SE MUESTRAN, SE VA A COLUMNAS Y SE AGREGAN LAS COLUMNAS TIPO DE 
       	'' FACTURACION
       	
       	'' este código de lee los datos de un fichero excel	
    	Dim filename As String
    	Dim sheetname As String
		Dim sheetname1 As String
    	
    	
		Me.dataGridView2.Enabled=True
    	Me.textBoxElegidos.Enabled=False
		Me.ButtonSepararDatos.Enabled=False

        ''filename = "C:\Users\BENI\Documents\Clientes\NACHO\SAGE\clientes.xls"
        filename = "C:\Users\anlla\OneDrive\Documentos\Clientes\NACHO\SAGE\clientes.xls"
        sheetname ="Clientes"
    	
    	
   	 If ((String.IsNullOrEmpty(fileName)) OrElse _
      (String.IsNullOrEmpty(sheetName))) Then _
      Throw New ArgumentNullException()

    Try
      Dim extension As String = IO.Path.GetExtension(fileName)

      Dim connString As String = "Data Source=" & fileName

      If (extension = ".xls") Then
        connString &= ";Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'"

      ElseIf (extension = ".xlsx") Then
        connString &= ";Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
      Else
        Throw New ArgumentException( _
          "La extensión " & extension & " del archivo no está permitida.")
      End If

      Using conexion As New System.Data.OleDb.OleDbConnection(connString)

	  ''Dim sql As String = "SELECT DISTINCT CLIENTE,NOMBRE,DIRECCIÓN,C.POSTAL,POBLACIÓN,PROVINCIA,EMAIL FROM [" & sheetName & "$]"
	  	Dim sql As String = "SELECT DISTINCT * FROM [" & sheetName & "$]"	  	
        Dim adaptador As New System.Data.OleDb.OleDbDataAdapter(sql, conexion)

        Dim dt As New DataTable("Excel")
        adaptador.Fill(dt)                
        conexion.Close()
        
        With DataGridViewDatos
        	.DataSource=dt
        End With
                
      End Using
      
	    Using conexion As New System.Data.OleDb.OleDbConnection(connString)

		Dim sql As String = "SELECT DISTINCT CDESCRIP,CCODIGO FROM [" & sheetName & "$] WHERE CCODIGO<>''"	 			
        Dim adaptador As New System.Data.OleDb.OleDbDataAdapter(sql, conexion)
        Dim dt As New DataTable("Excel")
        dt.Columns.Add("CDESCRIP")
        dt.Columns.Add("CCODIGO")
        adaptador.Fill(dt)
        conexion.Close()               
        With DataGridView2
        	.DataSource=dt
        	.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
        End With
      End Using
      
      
      DataGridViewDatos.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells

    Catch ex As Exception
    	MessageBox.Show(ex.Message.ToUpper,"error")
      Throw

    End Try
	
	'Me.buttonCargarDatos.Enabled=TRUE
	
   	DataGridView2.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
	
	
	
	Me.Cursor=System.Windows.Forms.Cursors.WaitCursor
	Try
			DataGridView1.Columns.Clear()
			
            For Each Col As DataGridViewColumn In DataGridViewDatos.Columns
            	DataGridView1.Columns.Add(DirectCast(Col.Clone, DataGridViewColumn))            
            Next
            
           ' MessageBox.Show(DataGridViewDatos.Rows.Count.ToString,"registros")
            For rowIndex As Integer = 0 To (DataGridViewDatos.Rows.Count - 1)
'            	MessageBox.Show(DataGridViewDatos.Rows(rowIndex).Cells("CIUDAD").Value.ToString)
    '                If instr(1,condicionfiltro,DataGridViewDatos.Rows(rowIndex).Cells("CCODGRUP").Value.ToString,1) Then           	
                    	DataGridView1.Rows.Add(DataGridViewDatos.Rows(rowIndex).Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray)
    '                End if	
				
            Next
            
			'' desactivo la ordenación por columnas            
            Dim i As Integer
            For i=0 To DataGridView1.Columns.Count-1
            	DataGridView1.Columns.Item(i).SortMode=DataGridViewColumnSortMode.NotSortable
            Next
	Catch ex As Exception
		MessageBox.Show(ex.Message.ToUpper,"error")
	End Try		
	Me.textBoxFiltrados.Text="Registros: " & DataGridView1.RowCount.ToString
	Me.Cursor=System.Windows.Forms.Cursors.Default

	'' no hay contacto
	haycontacto=False
       	
       	
    '' cargar los datos de un fichero dbase
    
    
    
'    	Me.DataGridView1.Columns.Clear()
'    
'    	Me.dataGridView2.Enabled=True
'    	Me.textBoxElegidos.Enabled=True
'    	Me.ButtonSepararDatos.Enabled=True
'    
'    	If Not(System.IO.Directory.Exists("c:\mailing")) Then
'    		'' no existe el directorio mailing, lo creo para trabajar
'    		System.IO.Directory.CreateDirectory("c:\mailing")	
'    	End If
'    	If Not(System.IO.Directory.Exists("c:\mailing\temporal")) Then
'    		'creo el directorio temporal
'    		System.IO.Directory.CreateDirectory("c:\mailing\temporal")	    		
'		End if    		
'		
'		'' limpio el directorio temporal
'		Try
'			For Each s As String In My.Computer.FileSystem.GetFiles("c:\mailing\temporal")
'	    	  My.Computer.FileSystem.DeleteFile(s)
'			Next
'		Catch e1 As Exception
'		End Try	
'		
'    	Dim origenfichero As string
'    	Dim origengrupos As String
'    	Dim origenprovincias As string
'    	
'    	origenfichero="c:\GrupoSP\GESTION\DBF02\Clientes.dbf"
'    	
'    	
'    	origengrupos="c:\GrupoSP\GESTION\DBF02\grupcli.dbf"
'    	
'    	origenprovincias="c:\GrupoSP\GESTION\DBF02\Provinc.dbf"
'    	
'    	If System.IO.File.Exists(origenfichero) Then
'    		System.IO.File.Copy(origenfichero,"c:\mailing\Clientes.dbf",True) '' copia el fichero y lo sobreescribe si existe...
'    	Else
'    		MessageBox.Show("No existe el fichero " & origenfichero,"Error...")
'    		Return
'    	End If
'    	
'    	If System.IO.File.Exists(origengrupos) Then
'    		System.IO.File.Copy(origengrupos,"c:\mailing\grupcli.dbf",True) '' copia el fichero y lo sobreescribe si existe...
'    	Else
'    		MessageBox.Show("No existe el fichero " & origengrupos,"Error...")
'    		Return
'    	End If
'    	
'    	If System.IO.File.Exists(origenprovincias) Then
'    		System.IO.File.Copy(origenprovincias,"c:\mailing\Provinc.dbf",True) '' copia el fichero y lo sobreescribe si existe...
'    	Else
'    		MessageBox.Show("No existe el fichero " & origenprovincias,"Error...")
'    		Return
'    	End If
'    	
    	
    	'' cargo los datos del fichero dbase
'    	Dim fi As New System.IO.FileInfo(origenfichero)
'    	
'    	Dim cn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=""dBase 5.0;text;CharacterSet=iso-8859-15"";Data Source='"  & fi.DirectoryName & "'")
'    	Dim TableName As String = fi.Name.Substring(0, fi.Name.Length - fi.Extension.Length)
'    	
'      	Dim cmd As New OleDb.OleDbCommand(TableName, cn)
'      	'cmd.CommandType = CommandType.TableDirect
'      	cmd.CommandText="SELECT DISTINCT CNOMCLI,CNOMCOM,CDIRCLI,CPOBCLI,CPTLCLI,CCONTACTO,EMAIL,CCODGRUP FROM " & TableName
'      	Dim dt As New DataTable
'      	cn.Open()
'      	Dim rdr As OleDb.OleDbDataReader = cmd.ExecuteReader
'      	dt.Load(rdr)
'      	DataGridViewDatos.DataSource = dt
'      	DataGridViewDatos.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells	
'      	
'      	cn.Close()
'      	cn.Dispose()
'      	      	
'      	''hay contacto
'      	
'      	haycontacto=True
'      	
'      	'' desactivo la ordenación por columnas            
'        Dim i As Integer
'        For i=0 To DataGridViewDatos.Columns.Count-1
'           	DataGridViewDatos.Columns.Item(i).SortMode=DataGridViewColumnSortMode.NotSortable
'        Next
'    	
'      	
'      	Dim fi1 As New System.IO.FileInfo(origengrupos)
'    	
'    	Dim cn1 As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase 5.0;Data Source='"  & fi1.DirectoryName & "'")    	
'    	TableName= fi1.Name.Substring(0, fi1.Name.Length - fi1.Extension.Length)
'    	
'      	Dim cmd1 As New OleDb.OleDbCommand(TableName, cn1)
'      	cmd1.CommandType = CommandType.TableDirect
'      	
'      	Dim dt1 As New DataTable
'      	cn1.Open()
'      	Dim rdr1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
'      	dt1.Load(rdr1)
'      	DataGridView2.DataSource = dt1
'      	DataGridView2.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
'    	
'    	cn1.Close()
'      	cn1.Dispose()
'      	
'      	
'      	Dim fi2 As New System.IO.FileInfo(origenprovincias)
'    	
'    	Dim cn2 As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=dBase 5.0;Data Source='"  & fi2.DirectoryName & "'")
'    	TableName= fi2.Name.Substring(0, fi2.Name.Length - fi2.Extension.Length)
'    	
'      	Dim cmd2 As New OleDb.OleDbCommand(TableName, cn2)
'      	cmd2.CommandType = CommandType.TableDirect
'      	
'      	Dim dt2 As New DataTable
'      	cn2.Open()
'      	Dim rdr2 As OleDb.OleDbDataReader = cmd2.ExecuteReader
'      	dt2.Load(rdr2)
'      	
'      	    	
'      	
'      	matrizprovincias=dt2.Select()
'      	
'      	
'      	
'      	For i=0 To matrizprovincias.getupperbound(0)
'      		matrizprovincias(i)("CNOMPROV")=matrizprovincias(i)("CNOMPROV").ToString.ToUpper
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"±","Ñ")
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"┴","Á")
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"ß","Á")
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"Ú","É")
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"Þ","É")
'      		
'      		
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"Ý","Í")
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"¾","Ó")
'      		
'      		matrizprovincias(i)("CNOMPROV")=Replace(matrizprovincias(i)("CNOMPROV"),"·","Ú")
'      		
'      		''MessageBox.Show(matrizprovincias(i)("CNOMPROV"),matrizprovincias(i)("CCODPROV"))
'      	Next
'      	
'    	cn2.Close()
'      	cn2.Dispose()
      	
      	    	
      	
      	Me.textBoxTotalRegistros.Text="Total registros: " & DataGridViewDatos.RowCount.ToString
      	
      	Me.ButtonSepararDatos.Enabled=True
      	
      	
      
    	
    '' este código de abajo lee los datos de un fichero excel	
    	'Dim filename As String
    '	Dim sheetname As String
    '	Dim ficherodb As String
    	
    	
    	
    '	ficherodb="C:\Users\BENI\Documents\nacho de la vega\Clientes.dbf"
    '	convertxls(ficherodb)
    '	filename="C:\Users\BENI\Documents\nacho de la vega\Clientes2.xls"	
    '	sheetname="Clientes"
    	
    'Me.buttonCargarDatos.Enabled=False	
    	
   	 'If ((String.IsNullOrEmpty(fileName)) OrElse _
     ' (String.IsNullOrEmpty(sheetName))) Then _
     ' Throw New ArgumentNullException()

    'Try
     ' Dim extension As String = IO.Path.GetExtension(fileName)

      'Dim connString As String = "Data Source=" & fileName

      'If (extension = ".xls") Then
      '  connString &= ";Provider=Microsoft.Jet.OLEDB.4.0;" & _
       '        "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'"

      'ElseIf (extension = ".xlsx") Then
      '  connString &= ";Provider=Microsoft.ACE.OLEDB.12.0;" & _
      '         "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
      'Else
      '  Throw New ArgumentException( _
      '    "La extensión " & extension & " del archivo no está permitida.")
      'End If

      'Using conexion As New System.Data.OleDb.OleDbConnection(connString)

		'Dim sql As String = "SELECT CNOMCLI,CNOMCOM,CDIRCLI,CPOBCLI,CPTLCLI,CCONTACTO,EMAIL,CCODGRUP FROM [" & sheetName & "$] WHERE CCODCLI<>'000001'"

		
       ' Dim adaptador As New System.Data.OleDb.OleDbDataAdapter(sql, conexion)

       ' Dim dt As New DataTable("Excel")

       ' adaptador.Fill(dt)
        
       ' conexion.Close()
        
       ' With DataGridViewDatos
       ' 	.DataSource=dt
       ' End With

      'End Using
      'DataGridViewDatos.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
	  ''dataGridViewDatos.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill)
    'Catch ex As Exception
     ' Throw

    'End Try
	
	'Me.buttonCargarDatos.Enabled=True

    End Sub
    
    Function convertxls(fn As String)
		Return true    	
	
End Function
    
    Sub DataGridViewDatosCellContentClick(sender As Object, e As DataGridViewCellEventArgs)
    	
    End Sub
    
    Sub Form1FormClosed(sender As Object, e As FormClosedEventArgs)
    	'Dim correcto As Integer=MessageBox.Show("Seguro que quieres salirr","Salir",MessageBoxButtons.YesNo,MessageBoxIcon.Question)    		
    	'	If correcto=DialogResult.Yes Then
    	'		Me.AxwebBrowser1.Dispose()
    	'		Me.AxwebBrowser1=Nothing		
    	'	Else	
    			'e.Cancel=true
		'		Return
    	'	End If
    End Sub
    
    Sub TextBoxMensajesTextChanged(sender As Object, e As EventArgs)
    	
    End Sub
    
    Sub DataGridView1CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
    	
    End Sub
    
    Sub DataGridView1RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
    	Me.textBoxFiltrados.Text="Registros: " & DataGridView1.RowCount.ToString
    End Sub
    
    Sub DataGridView2SelectionChanged(sender As Object, e As EventArgs)
    	'MessageBox.Show("cambiada seleccion","seleccion")    	
    	condicionfiltro="("
    	Me.textBoxElegidos.Text="Elegido(s): " & dataGridView2.SelectedRows.Count & vbCrLf    	
    	For Each row As DataGridViewRow In dataGridView2.SelectedRows
    		Me.textBoxElegidos.Text=Me.textBoxElegidos.Text & row.Cells("CCODIGO").Value.ToString.ToUpper & "-" & row.Cells("CDESCRIP").Value.ToString.ToUpper & vbcrlf	
    		condicionfiltro=condicionfiltro & row.Cells("CCODIGO").Value.ToString & ","  
    	Next
    	condicionfiltro=condicionfiltro & ")"
    	condicionfiltro=Replace(condicionfiltro,",)",")")
    	'Me.textBoxElegidos.Text=Me.textBoxElegidos.Text & condicionfiltro    	
    	'' dejo que el usuario pulse el botón de filtar...
    	''Me.ButtonSepararDatosClick(sender,e)
    End Sub
    
    Sub GroupBox1Enter(sender As Object, e As EventArgs)
    	'' nada por ahora
    End Sub
    
    Sub ButtonSalirClick(sender As Object, e As EventArgs)
		Me.Close()    	
    End Sub
    
    Sub Form1FormClosing(sender As Object, e As FormClosingEventArgs)
		Dim correcto As Integer=MessageBox.Show("Seguro que quieres salir","Salir",MessageBoxButtons.YesNo,MessageBoxIcon.Question)    		
    		If correcto=DialogResult.Yes Then
    			Me.AxwebBrowser1.Dispose()
    			Me.AxwebBrowser1=Nothing		
    		Else	
    			e.Cancel=true
				Return
    		End If    	
    End Sub
    
    Sub AxwebBrowser1DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
		'' nada por ahora    	
    End Sub
    
    Sub Form1Load(sender As Object, e As EventArgs)
		Me.ButtonCargarDatosClick(sender,e)
    End Sub
    
    Sub ButtonPararClick(sender As Object, e As EventArgs)
		pararproceso=true    	
    End Sub
    
    Sub DataGridView1RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
    	
    	'	Cells("CDESCRIP").Value.ToString    	
    If DataGridView1.Rows(e.RowIndex).Cells("CDIRCLI").Value.ToString.Length<=0 Or _
    	DataGridView1.Rows(e.RowIndex).Cells("CPOBCLI").Value.ToString.Length<=0 Or _
    	DataGridView1.Rows(e.RowIndex).Cells("CPTLCLI").Value.ToString.Length<=0 Then
		
     		Me.dataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor=Color.Red					
    Else    
			Me.dataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor=Color.black
    End If	
    End Sub
    
    
    
    
    Sub DataGridViewEnviosRowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
    '' no rula	
	'If DataGridViewEnvios.Rows(e.RowIndex).Cells(2).Value.ToString.Length<=0 Or _
   ' 	DataGridViewEnvios.Rows(e.RowIndex).Cells(3).Value.ToString.Length<=0 Or _
    '	DataGridViewEnvios.Rows(e.RowIndex).Cells(4).Value.ToString.Length<=0 Then
	'	
    ' 		Me.dataGridViewEnvios.Rows(e.RowIndex).DefaultCellStyle.ForeColor=Color.Red					
   ' Else    
	'		Me.dataGridViewEnvios.Rows(e.RowIndex).DefaultCellStyle.ForeColor=Color.black
   ' End If	    	
    End Sub
    
Sub EnviarEmail(emailenvio,fichero,snombre,sTituloEmail,stextoEmail)
    Dim _Message As New System.Net.Mail.MailMessage()

	Dim _SMTP As New System.Net.Mail.SmtpClient
	
	Dim tirahtml As String=File.ReadAllText(fichero,System.Text.Encoding.GetEncoding(1252))
	
	
	
	tirahtml=Replace(tirahtml,"file:///c:/firma/tiendainterior_fondo.jpg","cid:fondo")	
	tirahtml=Replace(tirahtml,"file:///c:/firma/logo.png","cid:logo")	
	tirahtml=Replace(tirahtml,"file:///c:/firma/firma.png","cid:firma")	
	tirahtml=Replace(tirahtml,"file:///c:/firma/facebook.png","cid:facebook")	
	tirahtml=Replace(tirahtml,"file:///c:/firma/twitter.jpg","cid:twitter")	
	tirahtml=Replace(tirahtml,"file:///c:/firma/pinterest.png","cid:pinterest")		
	tirahtml=Replace(tirahtml,"file:///c:/mailing/documentos/invitacion.jpg","cid:invitacion")		
	
	
	'' tema mensaje covid10
	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/crespon_negro.jpg","cid:crespon")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/logo.png","cid:logo")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/chon_nacho.png","cid:chon_nacho")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/facebook.png","cid:facebook")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/twitter.jpg","cid:twitter")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/pinterest.png","cid:pinterest")	
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/instagram-logo1.png","cid:instagram")
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/firma.png","cid:firma")
	'' se pasa a base64
	''tirahtml=Replace(tirahtml,"file:///c:/mailing/covid19/jpg/black-linen.png","cid:blacklinen")
			
				
	
	tirahtml=Replace(tirahtml,"[[NOMBRE]]",snombre)
	tirahtml=Replace(tirahtml,"[[EMAIL]]",emailenvio)
	tirahtml=Replace(tirahtml,"[[FECHA]]",Now.ToString("D"))
	tirahtml=Replace(tirahtml,"[[TEXTOEMAIL]]",stextoEmail)
	tirahtml=Replace(tirahtml,"á","&aacute;")
	tirahtml=Replace(tirahtml,"é","&eacute;")
	tirahtml=Replace(tirahtml,"í","&iacute;")
	tirahtml=Replace(tirahtml,"ó","&oacute;")
	tirahtml=Replace(tirahtml,"ú","&uacute;")
	tirahtml=Replace(tirahtml,"ñ","&ntilde;")
	tirahtml=Replace(tirahtml,"Ñ","&Ntilde;")
	tirahtml=Replace(tirahtml,"Á","&Aacute;")
	tirahtml=Replace(tirahtml,"É","&Eacute;")
	tirahtml=Replace(tirahtml,"Í","&Iacute;")
	tirahtml=Replace(tirahtml,"Ó","&Oacute;")
	tirahtml=Replace(tirahtml,"Ú","&Uacute;")
	
	
	Dim plainView As AlternateView = AlternateView.CreateAlternateViewFromString("Su cliente de e-mail no soportan html", Nothing,MediaTypeNames.Text.Plain)
	Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(tirahtml, Nothing, MediaTypeNames.Text.Html)

''Dim logo As New LinkedResource("c:\firma\logo.png",MediaTypeNames.Image.Jpeg)
	Dim logo As New LinkedResource("c:\mailing\covid19\jpg\logo.png",MediaTypeNames.Image.Jpeg)
	logo.ContentId="logo"
	htmlView.LinkedResources.Add(logo)	
	
	'Dim firma As New LinkedResource("c:\firma\firma.png",MediaTypeNames.Image.Jpeg)
	Dim firma As New LinkedResource("c:\mailing\covid19\jpg\firma.png",MediaTypeNames.Image.Jpeg)
	firma.ContentId="firma"
	htmlView.LinkedResources.Add(firma)
	'Dim facebook As New LinkedResource("c:\firma\facebook.png",MediaTypeNames.Image.Jpeg)
	Dim facebook As New LinkedResource("c:\mailing\covid19\jpg\facebook.png",MediaTypeNames.Image.Jpeg)
	facebook.ContentId="facebook"
	htmlView.LinkedResources.Add(facebook)
	'Dim twitter As New LinkedResource("c:\firma\twitter.jpg",MediaTypeNames.Image.Jpeg)
	Dim twitter As New LinkedResource("c:\mailing\covid19\jpg\twitter.jpg",MediaTypeNames.Image.Jpeg)
	twitter.ContentId="twitter"
	htmlView.LinkedResources.Add(twitter)
	'Dim pinterest As New LinkedResource("c:\firma\pinterest.png",MediaTypeNames.Image.Jpeg)	
	Dim pinterest As New LinkedResource("c:\mailing\covid19\jpg\pinterest.png",MediaTypeNames.Image.Jpeg)	
	pinterest.ContentId="pinterest"
	htmlView.LinkedResources.Add(pinterest)	
	
	'Dim instagram As New LinkedResource("c:\firma\instragram-logo1.png",MediaTypeNames.Image.Jpeg)	
	Dim instagram As New LinkedResource("c:\mailing\covid19\jpg\instagram-logo1.png",MediaTypeNames.Image.Jpeg)	
	instagram.ContentId="instagram"
	htmlView.LinkedResources.Add(instagram)	
	
	
	Dim chon_nacho As New LinkedResource("c:\mailing\covid19\jpg\chon_nacho.png",MediaTypeNames.Image.Jpeg)	
	chon_nacho.ContentId="chon_nacho"
	htmlView.LinkedResources.Add(chon_nacho)	
	
	
	Dim crespon As New LinkedResource("c:\mailing\covid19\jpg\crespon_negro.jpg",MediaTypeNames.Image.Jpeg)	
	crespon.ContentId="crespon"
	htmlView.LinkedResources.Add(crespon)	
	
	''Dim blacklinen As New LinkedResource("c:\mailing\covid19\jpg\black-linen.png",MediaTypeNames.Image.Jpeg)	
	''blacklinen.ContentId="blacklinen"
	''htmlView.LinkedResources.Add(blacklinen)	
	
	
	
	If System.IO.File.Exists("c:\mailing\documentos\invitacion.jpg") Then 
		''grafico de la invitacion
		Dim invitacion As New LinkedResource("c:\mailing\documentos\invitacion.jpg",MediaTypeNames.Image.Jpeg)	
		invitacion.ContentId="invitacion"
		htmlView.LinkedResources.Add(invitacion)
	End if
			
	_SMTP.Credentials = New System.Net.NetworkCredential("nachodelavega@nachodelavega.es", "NachoChon69#")

	_SMTP.Host = "mail.nachodelavega.es"

	_SMTP.Port = 25

	_SMTP.EnableSsl = False



'' quito espacios en blanco.
emailenvio=Replace(emailenvio," ","")
emailenvio=Replace(emailenvio,";",",")

_Message.From = New System.Net.Mail.MailAddress("info@nachodelavega.es", "Nacho de la Vega", System.Text.Encoding.UTF8) 'Quien lo envía
''_Message.From = New System.Net.Mail.MailAddress("info@nachodelavega.es", "Nacho de la Vega", System.Text.Encoding.UTF8) 'Quien lo envía
_Message.Subject = sTituloEmail

Dim replyto = New System.Net.Mail.MailAddress("nachodelavega@nachodelavega.es","Nacho de la Vega", System.Text.Encoding.UTF8) 'Quien recibe la respuesta



Try
	''_Message.[ReplyToList].Add(replyto)	
	_Message.To.Add(emailenvio) 'Cuenta de Correo al que se le quiere enviar el e-mail	
	_Message.ReplyToList.Add(replyto) '' cuenta de repilica 
	Catch ex As Exception
		''Throw  New Exception("Excption while trying to CreateEventSource for the EventLog: " & ex.Message)
		_Message.To.Add("beni@nachodelavega.es") 'Cuenta de Correo al que se le quiere enviar el e-mail	
		_Message.Subject = "Error enviando email a: " & emailenvio 
End Try



''_Message.From = New System.Net.Mail.MailAddress("nachodelavega@nachodelavega.es", "Nacho de la Vega", System.Text.Encoding.UTF8) 'Quien lo envía

''_Message.Subject = sTituloEmail

_Message.SubjectEncoding = System.Text.Encoding.UTF8 'Codificacion

'_Message.Body = tirahtml
_Message.AlternateViews.Add(plainView)
_Message.AlternateViews.Add(htmlView)

_Message.BodyEncoding = System.Text.Encoding.GetEncoding(1252)

_Message.Priority = System.Net.Mail.MailPriority.Normal

_Message.IsBodyHtml = True

If System.IO.File.Exists("c:\mailing\documentos\invitacion.jpg") Then
	Dim attachment As System.Net.Mail.Attachment	
	attachment = New System.Net.Mail.Attachment("c:\mailing\documentos\invitacion.jpg")
	_Message.Attachments.Add(attachment)
End If

Try

_SMTP.Send(_Message)
''MessageBox.Show("Enviado e-mail a: " & emailenvio)
_Message.Dispose()
Catch ex As System.Net.Mail.SmtpException	
	MessageBox.Show("Error enviando e-mail a: " & emailenvio & " " & ex.Message)
End Try

End Sub

    
    Sub ButtonPruebaClick(sender As Object, e As EventArgs)
    	If Me.checkBoxEmail.Checked Then
    		''  PARA VARIOS, SE SEPARAN POR COMAS
    		EnviarEmail(TextBoxEmailPrueba.Text.ToString(),sFichero,textBoxNombrePrueba.Text.ToString(),textBoxSubjectEmail.Text.ToString,textBoxEnvio.Text.ToString)    	
    		'EnviarEmail("beni@nachodelavega.es",sFichero,"PRUEBA ENVIO EMAIL",textBoxSubjectEmail.Text.ToString,textBoxEnvio.Text.ToString)    	
    		'EnviarEmail("majuca5@hotmail.com",sFichero,"MARIAN JUNQUERA",textBoxSubjectEmail.Text.ToString,textBoxEnvio.Text.ToString)    	
    		'EnviarEmail("ndelavega@hotmail.es",sFichero,"PRUEBA ENVIO EMAIL",textBoxSubjectEmail.Text.ToString,textBoxEnvio.Text.ToString)    	
    		'EnviarEmail("nachodelavega@nachodelavega.es",sFichero,"PRUEBA ENVIO EMAIL",textBoxSubjectEmail.Text.ToString,textBoxEnvio.Text.ToString)    	
    		MessageBox.Show("Email de prueba enviado a: beni@nachodelavega.es","E-MAIL")
		End IF    		
    End Sub
    
    Sub ButtonDatosProveedoresClick(sender As Object, e As EventArgs)
	'' este código de lee los datos de un fichero excel	
    	Dim filename As String
    	Dim sheetname As String
		Dim sheetname1 As String
    	
    	
		Me.dataGridView2.Enabled=False
    	Me.textBoxElegidos.Enabled=False
		Me.ButtonSepararDatos.Enabled=False    	
    	
    	filename="C:\mailing\xls\DIRECTORIO PROFESIONALES NACHODELAVEGA.xls"	
    	sheetname="Hoja2"
    	sheetname1="Hoja1"
    	
   	 If ((String.IsNullOrEmpty(fileName)) OrElse _
      (String.IsNullOrEmpty(sheetName))) Then _
      Throw New ArgumentNullException()

    Try
      Dim extension As String = IO.Path.GetExtension(fileName)

      Dim connString As String = "Data Source=" & fileName

      If (extension = ".xls") Then
        connString &= ";Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Extended Properties='Excel 8.0;HDR=YES;IMEX=1'"

      ElseIf (extension = ".xlsx") Then
        connString &= ";Provider=Microsoft.ACE.OLEDB.12.0;" & _
               "Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1'"
      Else
        Throw New ArgumentException( _
          "La extensión " & extension & " del archivo no está permitida.")
      End If

      Using conexion As New System.Data.OleDb.OleDbConnection(connString)

	  Dim sql As String = "SELECT DISTINCT CNOMCLI,OFICIO,CLIENTE,EMAIL,TELÉFONO,CDIRCLI,CPOBCLI,CPTLCLI,PROVINCIA FROM [" & sheetName & "$]"

	  sql=sql & "UNION SELECT DISTINCT CNOMCLI,OFICIO,CLIENTE,EMAIL,TELÉFONO,CDIRCLI,CPOBCLI,CPTLCLI,PROVINCIA FROM [" & sheetName1 & "$]"
        Dim adaptador As New System.Data.OleDb.OleDbDataAdapter(sql, conexion)

        Dim dt As New DataTable("Excel")

        adaptador.Fill(dt)
        
        conexion.Close()
        
        With DataGridViewDatos
        	.DataSource=dt
        End With

      End Using
      DataGridViewDatos.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
	  ''dataGridViewDatos.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill)
    Catch ex As Exception
    	MessageBox.Show(ex.Message.ToUpper,"error")
      Throw

    End Try
	
	'Me.buttonCargarDatos.Enabled=TRUE
	
	
	
	Me.Cursor=System.Windows.Forms.Cursors.WaitCursor
	Try
			DataGridView1.Columns.Clear()
			
            For Each Col As DataGridViewColumn In DataGridViewDatos.Columns
            	DataGridView1.Columns.Add(DirectCast(Col.Clone, DataGridViewColumn))            
            Next
            
           ' MessageBox.Show(DataGridViewDatos.Rows.Count.ToString,"registros")
            For rowIndex As Integer = 0 To (DataGridViewDatos.Rows.Count - 1)
'            	MessageBox.Show(DataGridViewDatos.Rows(rowIndex).Cells("CIUDAD").Value.ToString)
    '                If instr(1,condicionfiltro,DataGridViewDatos.Rows(rowIndex).Cells("CCODGRUP").Value.ToString,1) Then           	
                    	DataGridView1.Rows.Add(DataGridViewDatos.Rows(rowIndex).Cells.Cast(Of DataGridViewCell).Select(Function(c) c.Value).ToArray)
    '                End if	
				
            Next
            
			'' desactivo la ordenación por columnas            
            Dim i As Integer
            For i=0 To DataGridView1.Columns.Count-1
            	DataGridView1.Columns.Item(i).SortMode=DataGridViewColumnSortMode.NotSortable
            Next
	Catch ex As Exception
		MessageBox.Show(ex.Message.ToUpper,"error")
	End Try		
	Me.textBoxFiltrados.Text="Registros: " & DataGridView1.RowCount.ToString
	Me.Cursor=System.Windows.Forms.Cursors.Default

	'' no hay contacto
	haycontacto=False
	
	
    End Sub
    
     
    
    Sub ButtonGenerarEtiquetasClick(sender As Object, e As EventArgs)
    	
    	
       	Dim auxprov,nprovincia As String
    	Dim contaenvios As Integer=0    	    	
    	Dim nombre,contacto As String
		Dim direccion As String
		Dim poblacion As String
		Dim cpostal As String
		Dim email As String	
		Dim seleccionada As String
		Dim esrojo As Boolean
		
    	    
    	If Me.DataGridView1.Rows.Count<=0 Then
    	    	
    	    	MessageBox.Show("NO HAY DATOS CARGADOS, NO SE GENERAN ETIQUETAS","ERROR")
    	    	Return
		End If
		
		
		
			'' creo la conexión con la tabla y borro los registros
		Dim myCommand1 As New System.Data.OleDb.OleDbCommand
    	Dim MyConnection As System.Data.OleDb.OleDbConnection
    	MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\mailing\etiquetas\basedatos.mdb")    			    
		MyConnection.Open()
		Try			
    	    myCommand1.Connection=MyConnection
    	    myCommand1.CommandText="DELETE * FROM ETIQUETAS"
    	    myCommand1.ExecuteNonQuery()
		Catch ex As Exception
			MessageBox.Show(ex.Message.ToUpper,"error")
			Return
		End Try	
				
		
   	    Me.AxwebBrowser1.Visible=False
        
        '' añado las columnas
        Me.dataGridViewEnvios.Columns.Clear()        
        Me.dataGridViewEnvios.ColumnCount=5
        Me.dataGridViewEnvios.Columns(0).Name="Número"
        Me.dataGridViewEnvios.Columns(1).Name="Nombre"
        Me.dataGridViewEnvios.Columns(2).Name="Dirección"
        Me.dataGridViewEnvios.Columns(3).Name="Email"
        Me.dataGridViewEnvios.Columns(4).Name="Código Postal"
        Me.dataGridViewEnvios.Location=New Point (6,177)
        Me.dataGridViewEnvios.Visible=True
        Me.dataGridViewEnvios.Height=Me.AxwebBrowser1.Height
        Me.dataGridViewEnvios.Width=Me.AxwebBrowser1.Width
        Me.dataGridViewEnvios.AutoSizeColumnsMode=DataGridViewAutoSizeColumnMode.Fill
        'Me.DataGridViewEnvios.AutoSizeColumnsMode=DataGridViewAutoSizeColumnsMode.DisplayedCells
		
		Me.buttonSalir.Enabled=False
		Me.DataGridView1.Enabled=False
		
   		      
   		      'desactivo el botón de enviar
   		Me.buttonGenerarEtiquetas.Enabled=False      
    	Me.ButtonElegirFichero.Enabled=False
    	Me.ButtonCargarDatos.Enabled=False
    	Me.ButtonSepararDatos.Enabled=False
    	Me.buttonEnviar.Enabled=False
    	Me.buttonDatosProveedores.Enabled=False
    	    	
		Dim m As Object = System.Reflection.Missing.Value        
		
		Dim i As Integer
			
			
		pararproceso=False
		Me.buttonParar.Enabled=True
			
			
		Me.dataGridViewEnvios.Focus
			
					
		For Each row As DataGridViewRow In dataGridView1.Rows
				
				esrojo=False		
				seleccionada=row.Index
				
				'' para capturar los eventos...
				Application.DoEvents()
				If pararproceso Then					
					'' salgo del proceso
					Exit for		
				End If
				nprovincia=""
				
				If row.Cells("CPTLCLI").Value.ToString.Length=4 Then
					auxprov="0" & Strings.left(row.Cells("CPTLCLI").Value.ToString,1)
				Else
					auxprov=Strings.left(row.Cells("CPTLCLI").Value.ToString,2)					
				End If

				auxprov="00" & auxprov
			
				'' en el excel ya aparece el nombre de procincias
    	  		'For i=0 To matrizprovincias.getupperbound(0)
    	  		'	If matrizprovincias(i)("CCODPROV").ToString=auxprov Then
    	  		'		nprovincia=matrizprovincias(i)("CNOMPROV").ToString
    	  		'		
    	  		'	End If      		
    	  		'Next
    	  		
    	  		nprovincia=row.Cells("Provincia").Value.ToString.ToUpper

				nombre=row.Cells("CNOMCLI").Value.ToString.ToUpper
				nombre=Transforma(nombre)
				nombre=nombre.Replace("'","''")
				
				direccion=row.Cells("CDIRCLI").Value.ToString.ToUpper
				direccion=Transforma(direccion)
				direccion=direccion.Replace("'","''")
				
				poblacion=row.Cells("CPOBCLI").Value.ToString.ToUpper
				poblacion=Transforma(poblacion)
				poblacion=poblacion.Replace("'","''")
				
				cpostal=row.Cells("CPTLCLI").Value.ToString.ToUpper
				
				If cpostal.ToUpper.Length=4 Then
					cpostal="0" & cpostal
				End If
				
				'' si hay datos de contacto
				If haycontacto Then					
					contacto=row.Cells("CCONTACTO").Value.ToString.ToUpper
					contacto=Transforma(contacto)
					contacto=contacto.Replace("'","''")
				Else
					contacto=""
				End if	
				
				email=row.Cells("EMAIL").Value.ToString
				
				
				If direccion.Length<=0 Or poblacion.Length<=0 Or cpostal.Length<=0 Or nprovincia.length<=0 Then		        			
					esrojo=True
				Else
					Try
						myCommand1.CommandText="INSERT INTO ETIQUETAS (NOMBRE,DIRECCION,POBLACION,CODIGOPOSTAL,PROVINCIA) VALUES ('" & nombre & "','" & direccion & "','" _
							& poblacion & "','" & cpostal & "','" & nprovincia &"')"
    				    myCommand1.ExecuteNonQuery()
					Catch ex As Exception
						MessageBox.Show(ex.Message.ToUpper,"error")
						Return
					End Try	

		        End If
		        			    	    	    	   			   			
    	    	contaenvios=contaenvios+1
    	    	'Me.textBoxEnvio.Text=contaenvios.ToString & " -> Envio a " & row.Cells("CNOMCLI").Value.ToString  & vbCrLf & Me.textBoxEnvio.Text
    	    	dataGridViewEnvios.rows.Add(New String(){contaenvios.ToString, nombre, direccion,email,cpostal})
    	    	
    	    	If esrojo Then
    	    	'If direccion.Length<=0 or poblacion.Length<=0 or cpostal.Length<=0 or nprovincia.length<=0 Then		        	
    	    		Me.dataGridViewEnvios.Rows(Me.dataGridViewEnvios.RowCount-2).DefaultCellStyle.ForeColor=Color.Red
    	    	Else    	    	
    	    		Me.dataGridViewEnvios.Rows(Me.dataGridViewEnvios.RowCount-2).DefaultCellStyle.ForeColor=Color.black    	    	
        		End If
				
    	    	Me.DataGridViewEnvios.FirstDisplayedScrollingRowIndex = Me.DataGridViewEnvios.RowCount - 1
    	    	
       Next
       
       
        Me.buttonGenerarEtiquetas.Enabled=True
		Me.ButtonElegirFichero.Enabled=true
        Me.ButtonCargarDatos.Enabled=true
        Me.ButtonSepararDatos.Enabled=True
        Me.buttonSalir.Enabled=True
        Me.DataGridView1.Enabled=True
        Me.buttonDatosProveedores.Enabled=True

	   If pararproceso Then
	   		MessageBox.Show("Proceso interrumpido por el usuario","Fin de proceso")		   	
	        Return   		   
	   End If
	   
	   
	   
	   Me.buttonEspera.Width=Me.Width
	   Me.buttonEspera.Height=Me.Height
	   Me.buttonEspera.Top=0
	   Me.buttonEspera.Left=0
	   
	   Me.ButtonEspera.Location=New Point (6,177)
       Me.ButtonEspera.Height=Me.AxwebBrowser1.Height
       Me.ButtonEspera.Width=Me.AxwebBrowser1.Width
       
       
       Dim fnt As Font

	   fnt = Me.buttonEspera.Font

	   Me.buttonEspera.Font = New Font(fnt.Name, 50, FontStyle.Bold)
       
	   Me.buttonEspera.Visible=True
	  
	
    
    '' GENERAR ETIQUETAS Y MOSTRAR WORD CON LA INFORMACION	
    Try
	    Dim objWord As New Object
	    objWord = CreateObject("Word.Application") ' Creating a word application	    
		objWord.application.WindowState = 0 ' set the word window in normal state (Const wdWindowStatENormal = 0)
		''objWord.Documents.Open(System.AppDomain.CurrentDomain.BaseDirectory() & "\WORD.doc")
		''objWord.Documents.Open("c:\mailing\etiquetas\etiquetasvacias.doc")
		objWord.Documents.Open("c:\mailing\etiquetas\etiquetasvacias.doc")
		'objWord.ActiveDocument.MailMerge.OpenDataSource( _
		'Name:=System.AppDomain.CurrentDomain.BaseDirectory() & "ACCESSDB.mdb", _
		'Connection:="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ACCESSDB.mdb;Mode=Read;", _
		'SQLStatement:="select * from `TABLE_NAME`", _
	
	
	''Connection:="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ACCESSDB.mdb;Mode=Read;", _
		objWord.ActiveDocument.MailMerge.OpenDataSource( _
		Name:="c:\mailing\etiquetas\basedatos.mdb", _
		Connection:="Provider=Microsoft.Jet.OLEDB.4.0;;Mode=Read;", _
		SQLStatement:="select * from `ETIQUETAS`", _
		SubType:=Microsoft.Office.Interop.Word.WdMergeSubType.wdMergeSubTypeAccess)
		objWord.ActiveDocument.MailMerge.Execute()
		'objWord.ActiveDocument.SaveAs(System.AppDomain.CurrentDomain.BaseDirectory() & "\WORD1.DOC")
		'objWord.Documents.saveAs("c:\mailing\etiquetas\etiquetasUNO.doc")	
		objWord.Parent.Windows(2).Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
		Me.buttonEspera.Visible=False
		objWord.Visible = True	
	Catch ex As Exception
		MessageBox.Show(ex.Message.ToUpper,"error")
	End Try		
   
	
    End Sub
    
    
    End Class
        
    