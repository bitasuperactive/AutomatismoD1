#NoEnv ;// Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ;// Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ;// Ensures a consistent starting directory.
DetectHiddenWindows, On
OnExit("ExitFunc")
FormatTime, CurrentDate,, dd/MM/yy ;// Establecer formato de la fecha actual.

;----------------------------------------
; ROADMAP
;----------------------------------------
/*
*   1. Mejorar la verificación del Worksheet.
*   2. Implementar presentación de la app.
*/

;----------------------------------------
; Mensajes de texto informativos
;----------------------------------------
;// Presentación de la aplicación.
presentacion := "
(
No implementado
)"
;// *** MsgBox, 0, D1 Automatismo - Presentación, %presentacion%

;// Al abrir Internet Explorer, Edge lanza una ventana en blanco que opaca toda la pantalla.
info := "
(
Si al abrir Internet Explorer aparece una ventana en blanco opacando toda la pantalla, ve a:

Configuración de Microsoft Edge > Navegador predeterminado
> Permitir que Internet Explorer abra sitios en Microsoft Edge

Y establece este parámetro a <Nunca>.
)"
MsgBox, 0, D1 Automatismo - Información, %info%

;// A falta de la función "OnError()" en esta versión de AutoHotkey,
;// se implementa un bloque try/catch para todo el código simulando su funcionalidad.
try
{
    ;----------------------------------------
    ; Columnas Excel objetivo
    ;----------------------------------------
    global fechaSolicitudCol := "F"
    global cupsCol := "G"
    global comerCol := "H"
    global distriCol := "I"
    global excedentesCol := "K"
    global colectivoCol := "M"
    global tipoInstCol := "N"
    global idClienteCol := "P"
    global tipoAutoconsumoCol := "R"
    global contactoCol := "Z"
    global fechaContactoCol := "AA"
    global estadoContactoCol := "AB"
    global observacionesCol := "AH"

    ;----------------------------------------
    ; Excel
    ;----------------------------------------
    ;// Solicitar al usuario que seleccione el Workbook a gestionar.
    FileSelectFile, path, 1,, Libro de excel (*.xlsx)
    ;// Cerrar la aplicación si no se selecciona una ruta válida.
    If (!path) {
        MsgBox, 0, Automatismo D1 - ERROR, No se ha seleccionado ningún Workbook. La aplicación se cerrará.
        ExitApp, -1
    }
    ;// Crear una instancia de Excel.
    global excel := ComObjCreate("Excel.Application")
    global excelHWND := excel.Hwnd
    excel.Visible := True
    WinWait, % "ahk_id " excelHWND
    WinMaximize
    ;// Abrir el Workbook seleccionado por el usuario y activar el primer Worksheet.
    global workbook := excel.Workbooks.Open(path)
    workbook.Sheets(1).Activate()
    ;// Verificación mínima de la estructura del WorkSheet.
    While, excel.Range(cupsCol . "1").Value != "CUPS" or excel.Range(contactoCol . "1").Value != "Contacto"
    {
        MsgBox, 1, Automatismo D1, La estructura del excel es errónea, corrigela posicionando el encabezado "CUPS" en la columna "G" y, el encabezado "Contacto"
        IfMsgBox Cancel
        throw Exception("Estructura del WorkSheet (1) inválida.", -1)
    }

    ;----------------------------------------
    ; Internet Explorer
    ;----------------------------------------
    ;// Crear una instancia de Internet Explorer.
    global ie := ComObjCreate("InternetExplorer.Application")
    global ieHWND := ie.Hwnd
    global frame
    ComObjConnect(ie, IE_Events)
    ie.Visible := True
    WinWait, % "ahk_id " ieHWND
    WinActivate
    Send, #{Right}
    ;// Navegar a IC Web.
    global crmUrl := "CRM_URL"
    ie.Navigate(crmUrl)
    ;// *** Si se excede el tiempo de espera o,
    ;// no se encuentra el cuadro de entrada "login", lanzar excepción.
    If (!ObjBusy(ie) or !ie.Document.getElementById("login")) {
        throw Exception("CRM no responde: Tiempo de espera excedido.", -1)
    }
    ;// Iniciar sesión.
    username := ie.Document.getElementById("login")
    password := ie.Document.getElementById("pass")
    submit := ie.Document.getElementById("bentrar")
    username.Value := "USERNAME"
    password.Value := "PASSWORD"
    submit.Click()
    ;// *** Si se excede el tiempo de espera o, el campo de nombre de usuario
    ;// continúa presente tras la carga del DOM, lanzar excepción.
    If (!ObjBusy(ie) or !GetFrame()) {
        throw Exception("CRM no responde: Tiempo de espera excedido.", -1)
    }
    Else If (ie.Document.getElementById("login")) {
        throw Exception("Nombre o clave de acceso incorrectos.", -1)
    }

    ;----------------------------------------
    ; Anotaciones
    ;----------------------------------------
    ;// Endesa Andalucía con excedentes, individual.
    an0023Ind := "
(
Recibida solicitud Autoconsumo del distribuidor para el CUPS EX00000000000000000000.
Tipo Autoconsumo CON EXCEDENTES, INDIVIDUAL.
Documentación:
- CIE generación
- CIEBT tramitado a través de PUES y Justificante de presentación de documentación
- Autocontrato
Si el cliente no tiene acceso a la web, indicar dirección para el envío del autocontrato.
Por favor, confirmar todo lo anterior y, en caso de ser correcto, grabar consentimiento para poder continuar con la gestión.
)"
    ;// Endesa Cataluña con excedentes, individual.
    anEndesaCatalunaInd := "
(
Recibida solicitud Autoconsumo del distribuidor para el CUPS EX00000000000000000000.
Tipo Autoconsumo CON EXCEDENTES, INDIVIDUAL.
Documentación:
- Autocontrato
Si el cliente no tiene acceso a la web, indicar dirección para el envío del autocontrato.
Por favor, confirmar todo lo anterior y, en caso de ser correcto, grabar consentimiento para poder continuar con la gestión.
)"
;// Cualquier distribuidora, individual.
    an0000Ind := "
(
Recibida solicitud Autoconsumo del distribuidor para el CUPS EX00000000000000000000.
Tipo Autoconsumo XXX EXCEDENTES, INDIVIDUAL.
Documentación:
- CIE generación
- Autocontrato
Si el cliente no tiene acceso a la web, indicar dirección para el envío del autocontrato.
Por favor, confirmar todo lo anterior y, en caso de ser correcto, grabar consentimiento para poder continuar con la gestión.
)"
    ;// Cualquier distribuidora, colectivo.
    an0000Сol := "
(
Recibida solicitud Autoconsumo del distribuidor para el CUPS EX00000000000000000000.
Tipo Autoconsumo XXX EXCEDENTES, COLECTIVO.
Documentación:
- CIE generación
- Acuerdo de reparto en formato PDF con la firma de todos los participantes en el autoconsumo colectivo.
- Fichero generado a partir del acuerdo de reparto
- Autocontrato
Si el cliente no tiene acceso a la web, indicar dirección para el envío del autocontrato.
Al tratarse de un suministro COLECTIVO, confirmar con el cliente si se trata de constante o variable, y en caso de ser constante, que confirme el coeficiente
Por favor, confirmar todo lo anterior y, en caso de ser correcto, grabar consentimiento para poder continuar con la gestión.
)"

    ;----------------------------------------
    ; Tramitar contactos
    ;----------------------------------------
    ;// Seguimiento del número de solicitudes a gestionar.
    global taskCount := 0
    ;// Seguimiento de los contactos creados y obtenidos para evitar duplicados.
    global contacts := []
    ;// Iterar sobre todas las solicitudes del Worksheet.
    While, excel.Range("A" . A_Index).Value != ""
    {
        If (A_Index = 1) {
            ;// Desactivar la actualización de la interfaz de Excel.
            Excel_Toggle_Screen_Update(1)
            ;// Contabilizar solicitudes a gestionar.
            CountTasks()
            ;// Remplazar los puntos por barras en las fechas de solicitud para su directa introducción
            ;// en los formularios de los contactos.
            excel.Range(fechaSolicitudCol . ":" . fechaSolicitudCol).Replace(What:=".", Replacement:="/")
            ;// Saltar los encabezados.
            Continue
        }

        ;// Restaurar el color de la fila anterior.
        If (A_Index > 2) {
            excel.Range((A_Index - 1) . ":" . (A_Index 1)).Interior.ColorIndex := 0
        }
        ;// Señalar la fila a tratar en color amarillo.
        excel.Range(A_Index . ":" . A_Index).Interior.ColorIndex := 6

        ;// Reduce el consumo de recursos de Excel, actualizando su IGU únicamente en este punto del programa.
        Excel_Toggle_Screen_Update(2)

        ;----------------------------------------
        ; Establecer variables a tratar.
        ;----------------------------------------
        global cups := excel.Range(cupsCol . A_Index).Value
        global distri := excel.Range(distriCol . A_Index).Value
        global comer := excel.Range(comerCol . A_Index).Value
        global tipoInst := excel.Range(tipoInstCol . A_Index).Value
        global idCliente := excel.Range(idClienteCol . A_Index).Value
        global tipoAutoconsumoCelda := excel.Range(tipoAutoconsumoCol . A_Index)
        global contactoCelda := excel.Range(contactoCol . A_Index)
        global fechaContactoCelda := excel.Range(fechaContactoCol . A_Index)
        global estadoContactoCelda := excel.Range(estadoContactoCol . A_Index)
        global observacionesCelda := excel.Range(observacionesCol . A_Index)

        colectivo := (excel.Range(colectivoCol . A_Index).Value = "SI") ? True : False
        excedentes := (excel.Range(excedentesCol . A_Index).Value = "CON EXCEDENTES") ? True : False
        descripcion := (distri = 24 or distri = 359 or distri = 396) ? "AUTOCONSUMO CATALUÑA" : "AUTOCONSUMO"
        anotaciones := ""

        ;// Esta versión de AutoHotkey no admite la función <Switch>.
        If (distri = 23 and !colectivo and excedentes)
            anotaciones := an0023Ind
        Else If ((distri = 24 or distri = 359 or distri = 396) and (!colectivo and excedentes))
            anotaciones := anEndesaCatalunaInd
        Else {
            anotaciones := (colectivo) ? an0000Col : an0000Ind
            anotaciones := (excedentes) ? StrReplace(anotaciones, "XXX", "CON") : StrReplace(anotaciones, "XXX", "SIN")
        }

        ;// Introducir el CUPS del cliente en las anotaciones.
        anotaciones := StrReplace(anotaciones, "EX00000000000000000000", cups)

        ;// Si ya consta contacto u observación en Excel, continuar con la siguiente solicitud.
        ;// Omitir los valores clave '0' y "FORCE" de las observaciones.
        If (contactoCelda.Value != "" or (observacionesCelda.Value != "" and observacionesCelda.Value != 0 and observacionesCelda.Value != "FORCE")) {
            Continue
        }

        ;----------------------------------------
        ; Determinar el tipo de autoconsumo
        ;----------------------------------------
        If (!colectivo and !excedentes and tipoInst = "Red interior")
            tipoAutoconsumoCelda.Value := 31
        Else if (colectivo and !excedentes and tipoInst = "Red-varios consumidores")
            tipoAutoconsumoCelda.Value := 32
        Else If (!colectivo and excedentes and tipoInst = "Red interior")
            tipoAutoconsumoCelda.Value := 41
        Else If (colectivo and excedentes and tipoInst = "Red-varios consumidores")
            tipoAutoconsumoCelda.Value := 42
        Else If (colectivo and excedentes and tipoInst = "A través de red")
            tipoAutoconsumoCelda.Value := 43

        ;----------------------------------------
        ; Buscar al cliente
        ;----------------------------------------
        If (!SearchClient())
            Continue

        ;----------------------------------------
        ; Leer contactos abiertos
        ;----------------------------------------
        ;// Si la pestaña "Contactos" se encuentra resaltada, entrar.
        ;// Si la palabra clave "FORCE" es indicada en las obversaciones, omitir esta comprobación.
        contactos := frame.getElementById("ICWEB-itm-1-txt")
        If (contactos.getAttribute("style") = "background-color: orange;" and observacionesCelda.Value != "FORCE") {
			;// Obtener id del punto de suministro.
			idSuministro := ObtainSupplyPointId()
			;// Entrar en la pestaña de Contactos.
			GoToTab(2)
			;// Seleccionar punto de suministro.
			If (!SelectOption(ie, frame, "psuministro", idSuministro)) {
				observacionesCelda.Value := "INTERVENCIÓN MANUAL REQUERIDA: No ha sido posible seleccionar el punto de suministro del cliente para leer sus contactos."
				;// Volver a la pestaña "Localización".
				GoToTab(1)
				Continue
			}
			;// Buscar contactos abiertos de Autoconsumo.
			readContactsResult := ReadContacts()
			If (readContactsResult = 2)
				Continue
			;// Volver a la pestaña "Localización".
			GoToTab(1)
        }

        ;----------------------------------------
        ; Rellenar formulario del contacto
        ;----------------------------------------
        register := frame.getElementById("registrar")
        register.Click()
        ObjBusy(ie)

        ;// Obtener el objeto del formulario resultante, si no ha sido obtenido.
        If (!IsObject(form)) {
            form := IEGet("Registro de Contactos")
            global formHWND := form.Hwnd
            WinWait, % "ahk_id " formHWND
            WinActivate
            Send, #{Left}
        }

        ObjBusy (form)

        ;// Seleccionar sociedad (4=COR, 1=ML).
        targetOption := (comer = 642) ? "4" : "1"

        SelectOption (form, form.Document, "select_Sociedad", targetOption)
        ;// Seleccionar contrato.
        If (!SelectOption(form, form.Document, "select_Contractual", cups)) {
            observacionesCelda.Value := "INTERVENCIÓN MANUAL REQUERIDA: No ha sido posible seleccionar el CUPS al crear el contacto"
            Continue
        }
        ;// Seleccionar canal: Distribuidora.
        SelectOption (form, form.Document, "selectCanal", "A00")
        ;// Seleccionar reclamante: Titular.
        SelectOption(form, form.Document, "selectORG", "000")
        ;// Rellenar descripción.
        form.Document.getElementById("descripcionContacto").innerText := descripcion
        ;// Seleccionar tipo de acción: Gestión.
        SelectOption(form, form.Document, "selectN1TipoAccion", "A0")
        ;// Seleccionar área: Ventas y contratación.
        SelectOption(form, form.Document, "selectN2SubCat", "A0-ABC00000")
        ;// Seleccionar cuestión: Autoconsumo.
        SelectOption(form, form.Document, "selectN3Cuestion", "A0-ABC00000-A000")
        ;// Rellenar anotaciones.
        form.Document.getElementById("textAreaContactos").textContent := anotaciones
        ;// Si no hay respuesta seleccionada o el valor indicado no es válido,
        ;// asumir que no consta medio de contacto y marcar "No necesaria respuesta".
        emailChecked := form.Document.getElementById("check_email_cliente").hasAttribute("checked")
        emaillistBox := form.Document.getElementById("select_Email")
        emailValue := emailListBox.getElementsByTagName("option")[emailListBox.selectedIndex].getAttribute("value")
        smsChecked := form.Document.getElementById("check_tlfn_cliente").hasAttribute("checked")
        smsListBox := form.Document.getElementById("selectMovil")
        smsValue := smsListBox.getElementsByTagName("option")[smsListBox.selectedIndex].getAttribute("value")
        otroChecked := form.Document.getElementById("check_tlfn_otro").hasAttribute("checked")
        otroListBox := form.Document.getElementById("selectotro")
        otroValue := otroListBox.getElementsByTagName("option")[otroListBox.selectedIndex].getAttribute("value")
        If ((emailChecked and !InStr(emailValue, "@")) or (smsChecked and !InStr(smsValue, "6")) or (otroChecked and !InStr(otroValue, "6") and !InStr(otroValue, "9"))) {
            form.Document.getElementById("respuesta_no").checked := True
        }
        ;// Escalar contacto.
        escalar := form.Document.getElementById("chkEscalar-img")
        escalar.Click()
        ObjBusy(form)
        ;// Departamento: BO Gestión Autoconsumo (Endesa Cataluña) ó Atención Cliente Comercializadora.
        targetOption := (distri = 24 or distri = 359 or distri = 396) ? "988" : "097"
        SelectOption(form, form.Document, "selectDepartamento", targetOption)
        ObjBusy(form)

        ;// Grabar contacto.
        MsgBox, 1, Automatismo D1 - PAUSA DE SEGURIDAD, % "AL PULSAR ACEPTAR PARA GRABAR EL CONTACTO.`nPULSA CANCELAR PARA OMITIR LA GRABACIÓN."
        IfMsgBox, Cancel
        {
            Continue
        }
        grabar := form.Document.getElementById("bgrabar")
        grabar.Click()
        ObjBusy(form)
        ;// Obtener contacto.
        output := form.Document.getElementById("contenido-mensaje-servidor").getElementsByTagName("p")[0].innerText
        targetPos := InStr(output, "6")
        If (targetPos = 0) {
            ;// Contacto no obtenido.
            observacionesCelda.Value := "ERROR: " . output
        }
        Else {
            ;// Contacto obtenido.
            SetContact(Substr(output, targetPos, 10), CurrentDate, "1er envío cliente ", "")
        }
    }

    ;----------------------------------------
    ; Fin de la aplicación
    ;----------------------------------------
    ;// Restrablecer el color de la fila tratada.
    excel.UsedRange.Interior.ColorIndex := 0
    Log("")
    ExitApp
}
catch e
{
	Log(e)
	ExitApp
}

;----------------------------------------
; Funciones
;----------------------------------------
;// Registrar el resultado de las gestiones y los datos de la excepción, si la hubiera, en un archivo de texto.
Log(exception)
{
	FormatTime, CurrentDateTime,, dd/MM/yy HH:mm:ss
	FormatTime, CurrentDate,, ddMMyy

	contactsCount := contacts.MaxIndex()
	tasksLeft := taskCount - contactsCount
	error = (exception) ? "Error on line " . exception.Line . ": " . exception.Message "`n`n`n" : " "

	message = 
(
[%CurrentDateTime%]
Solicitudes a gestionar: %taskCount%
Contactos creados/obtenidos: %contactsCount%
Solicitudes pendientes de intervención manual: %tasksLeft%

%error%
)

	FileAppend % message, [ATR][Autoconsumo]LogRobot_%CurrentDate%.txt
	Return True
}

ExitFunc(ExitReason, ExitCode)
{
	;// Actualizar la interfaz de Excel.
	Excel_Toggle_Screen_Update(1)

	;// *** Guardar Workbook.
	try workbook.Save()
	catch e
		Log(e)

	;// Cerrar ventanas abiertas por el programa.
	If (WinExist("ahk_id " ieHWND))
		WinClose
	If (WinExist("ahk_id " formHWND))
		WinClose
	;If (WinExist("ahk_id " excelHWND))
	;	WinClose
}

;// Alternar la actualización de la IGU de Excel.
Excel_Toggle_Screen_Update(n)
{
	Loop, n
	{
		excel.Application.ScreenUpdating := !excel.Application.ScreenUpdating
		Sleep, 50
	}
}

;// Contabilizar la solicitudes a gestionar.
CountTasks()
{
	While, excel.Range("A" . (i := A_Index + 1)).Value != ""
	{
		contactoValue := excel.Range(contactoCol . i).Value
		observacionesValue := excel.Range(observacionesCol . i).Value
		If (contactoValue = "" and (observacionesValue = "" or observacionesValue = 0 or observacionesValue = "FORCE"))
			taskCount++
	}
}

;// Establecer el tipo de autoconsumo correspondiente a cada solicitud.
SetSelfconsumptionTypes()
{
	While, excel.Range("A" . (i := A_Index + 1)).Value != ""
	{
		colectivo := (excel.Range(colectivoCol . i).Value = "SI") ? True : False
		excedentes := (excel.Range(excedentesCol . i).Value = "CON EXCEDENTES") ? True : False
		tipoInst := excel.Range(tipoInstCol . i).Value
		tipoAutoconsumoCelda := excel.Range(tipoAutoconsumoCol . i)
		If (!colectivo and !excedentes and tipoInst = "Red interior")
			tipoAutoconsumoCelda.Value := 31
		Else if (colectivo and !excedentes and tipoInst = "Red-varios consumidores")
			tipoAutoconsumoCelda.Value := 32
		Else If (!colectivo and excedentes and tipoInst = "Red interior")
			tipoAutoconsumoCelda.Value := 41
		Else if (colectivo and excedentes and tipoInst = "Red-varios consumidores")
			tipoAutoconsumoCelda.Value := 42
		Else if (colectivo and excedentes and tipoInst = "A través de red") 
			tipoAutoconsumoCelda.Value := 43
	}
}

;// Control de los eventos generados por Internet Explorer.
class IE_Events
{
	OnQuit()
	{
		If A_ExitReason not in Menu,Exit
			Log(Exception("La aplicación Internet Explorer se ha cerrado inesperadamente.", -1))
	}
}

;// Esperar a que el objeto se encuentre disponible, con un tiempo límite de 10.000 milisegundos (20 iteraciones).
;// Devuelve verdadero si el objeto no ha excedido el tiempo límite, falso en su defecto.
ObjBusy(object)
{
	While, object.Busy or object.ReadyState != 4 or !object.Document or object.Document.ReadyState != "complete"
	{
		If (A_Index >= 20)
			return False

		Sleep, 500
	}
	return True
}

;// Establece y devuelve el DOM del frame objetivo.
GetFrame()
{
	return frame := ie.Document.ParentWindow.Frames[1].Document
}

;// Obtener objecto preexistente de Internet Explorer correspondiente al nombre de la ventana indicado.
IEGet(Name)
{
	IfEqual, Name,, WinGetTitle, Name, ahk_class IEFrame
	Name := (Name = "Consulta - HC - Windows Internet Explorer") ? "about:Tabs" : RegExReplace(Name, " - (Windows|Microsoft)? ?Internet Explorer$")
	for wb in ComObjCreate("Shell.Application").Windows()
		if wb.LocationName = Name and InStr(wb.FullName, "iexplore.exe")
			return wb
}

;// Buscar cliente.
;// Devuelve Verdadero si ha sido posible obtener al cliente, Falso en su defecto.
SearchClient()
{
	ObjBusy(ie)
	GetFrame()

	;// Si el botón "Buscar" no se encuentra, realizar nueva búsqueda.
	If (searchOther := frame.getElementById("buscar_otro")) {
		searchOther.Click()
		ObjBusy(ie)
		GetFrame()
	}

	;// Buscar por CUPS para obtener el id del punto de suministro
	;// a menos que se indique el caracter especial en observaciones, en cuyo caso se busca por IC.
	If (observacionesCelda.Value != "FORCE") {
		frame.getElementById("psuminis").checked := True
		ObjBusy(ie)
		inputCups := frame.getElementById("psuminis_cups")
		inputCups.Value := cups
	}
	Else {
		inputIc := frame.getElementById("cliente_ic_cliente")
		inputIc.Value := idCliente
	}
	submit := frame.getElementById("buscar")
	submit.Click()

	;// Si se excede el tiempo de espera, restablecer CRM y continuar la iteración del bucle.
	;// Refrescar CRM sin reenviar el formulario.
	If (!ObjBusy(ie)) {
		If (WinExist("Mensaje de página web")) {
			observacionesCelda.Value := "NO EXISTEN INTERLOCUTORES COMERCIALES"
			WinClose
		}
		Else {
			observacionesCelda.Value := "TIEMPO DE ESPERA EXCEDIDO"
			ie.Stop()
			ie.Navigate(crmUrl)
		}
		ObjBusy(ie)
		return False
	}

	GetFrame()

	;// Si el cliente tiene un bloqueo automático para EDP, continuar la iteración el bucle.
	If (frame.querySelector("#imita-popup > div.popup-esq-der-bottom > div > div > div.contenido-popup > div > input[type=button]:nth-child(3)")) { 
		MsgBox, 0, Automatismo D1, Cliente con bloqueo automático para EDP.
		observacionesCelda.Value := "CLIENTE CON BLOQUEO AUTOMÁTICO PARA EDP"
		return False
	}

	idTextBox := frame.querySelector("#cliente_ic_cliente")
	idObtenido := idTextBox.getAttribute("value")
	;// Comprobar que se ha obtenido algún IC.
	If (idObtenido = "") {
		observacionesCelda.Value := "CRM NO HA DEVUELTO DATOS PARA ESTE IC"
		return False
	}

	;// Comprobar que el IC resultante equivale al indicado en el excel.
	If (idCliente and idCliente != idObtenido) {
		observacionesCelda.Value := "CUPS NO COINCIDENTE CON EL IC DEL CLIENTE"
		return False
	}

	return True
}

;// Obtener el id del punto de suministro correspondiende a la solicitud.
ObtainSupplyPointId()
{
	return Substr(frame.getElementById("treeResultados3 -cnt-start").textContent, 3, 10)
}

;// Se desplaza a la pestaña indicada de CRM.
;// 1 = Localización, 2 = Contactos.
GoToTab(tab)
{
	ObjBusy(ie)
	GetFrame()
	localizacionTab := frame.getElementById("ICWEB-itm-0")
	contactosTab := frame.getElementById("ICWEB-itm-1-txt")
	If (tab = 1)
		localizacionTab.Click()
	Else If (tab = 2)
		contactosTab.click()
	;*** Else Throw Exception

	ObjBusy (ie)
	GetFrame()
}

;// Lee los contactos abiertos del cliente, si los hubiera.
;// Esta versión de AutoHotKey carece de la instrucción <do while>.
;// 0 = No existen contactos abiertos para este CUPS.
;// 1 = Existen contactos abiertos pero no han sido reconocidos como autoconsumo.
;// 2 = Existen contactos abiertos que sí han sido reconocidos como autoconsumo.
ReadContacts()
{
	ObjBusy(ie)
	GetFrame()

	openedContacts := False
	firstIteration := True
	sociedad := (comer = 642) ? "COR" : "ML"

	;// Itera sobre toda la lista de contactos.
	While, firstIteration or ((downButton := frame.getElementById("tvContactos_pager-btn-4")) and downButton.hasAttribute("dsbl")) { 
		
		If (firstIteration)
			firstIteration:= False
		Else
			downButton.Click()

		ObjBusy(ie)
		GetFrame()

		tvContactos := frame.getElementById("tvContactos")
		table := tvContactos.getElementsByTagName("table")[0]
		rows := table.getElementsByTagName("tr")

		;// Si no existen contactos para el cliente, devolver 0.
		If (rows[1].getAttribute("rr") = 0)
			return 0

		While, (i := A_Index) < rows.Length {
			
			data := rows[i].getElementsByTagName("td")
			;// Si existe un contacto que no esté anulado ni cerrado...
			condition1 := data[3].textContent != "Anulado" and data[3].textContent != "Cerrado"
			If (condition1) {
				openedContacts := True
				;// y, que sea de "Ventas y Contratación" y "Autoconsumo" o, cuyas anotaciones hagan referencia al autoconsumo...
				condition2 := data[7].textContent = sociedad and data[9].textContent = "Ventas y Contratación" and data[10].textContent = "Autoconsumo"
				condition3 := (condition2) ? False : CheckAnotations(data[1].getElementsByTagName("a")[0], ["autoconsumo", "d1", "placas solares"])
				If (condition2 or condition3) {
					condition4 := CheckAnotations(data[1].getElementsByTagName("a")[0], ["Se recibe documentación"])
					;// introducirlo en el excel, junto a la fecha de creación.
					state := (condition4) ? "Documentación recibida " : "1er envío cliente "
					SetContact(data[1].textContent, FormatDate(data[4].textContent), state, "YA CONSTABA CREADO")
					;// Volver a la pestaña "Localización".
					GoToTab (1)
					ObjBusy(ie)
					return 2
				}
			}
		}
	}
	
	return (openedContacts) ? 1 : 0
}

;// Busca palabras clave en las anotaciones de los contactos.
;// Verdadero Se ha encontrado alguna coincidencia, Falso en su defecto.
CheckAnotations(link, needles)
{
	link.Click()

	;// *** La función ObjBusy(ie) no funciona correctamente aqui.
	While, !anotationDiv or !(anotations := anotationDiv.textContent)
	{
		Sleep, 50
		GetFrame()
		anotationDiv := frame.querySelector("#imita-popup > div.popup-esq-der-bottom > div > div > div.contenido-popup > div > pre")
	}

	closeButton := frame.querySelector("#a_vent_cerrar")
	closeButton.Click()

	While, (i = A_Index) <= needles.MaxIndex()
		If (InStr(anotations, needles[i]))
			return True

	;// *** Seguimiento de las anotaciones no reconocidas.
	MsgBox, Anotaciones no encontradas.

	return False
}

;// Transforma una fecha en el formato dd/MM/yyyy.
FormatDate(date)
{
	dateArray := StrSplit(Substr(date, 1, 10), ".")
	return dateArray[1] . "/" . dateArray[2] . "/" . dateArray[3]
}

;// Seleccionar una opción desde un objeto ListBox.
;// Devuelve Falso si no existe la opción objetivo dentro del ListBox.
SelectOption(object, dom, id, targetOption)
{
	;// Si la opción objetivo es nula, devolver Falso.
	If (!targetOption)
		return False

	listBox := dom.getElementById(id)
	options := listBox.getElementsByTagName("option")
	While, (i = A_Index) < options . Length
		If (InStr(options[i].getAttribute("value"), targetOption)) {
			listBox.selectedIndex := i
			try listBox.onchange()
			ObjBusy(object)
			return True
		}

	return False
}

;// Introduce el contacto en el Excel, comprobando que no sea un duplicado.
;// Devuelve Verdadero si la inserción ha sido satisfactoria, Falso en su defecto.
SetContact(contact, date, state, observations)
{
	While, (i := A_Index) <= contacts.MaxIndex()
		If (contacts[i] = contact) {
			observacionesCelda.Value := "CONTACTO DUPLICADO: " . contact
			return False
		}

	contacts.Push(contact)

	contactoCelda.Value := contact
	fechaContactoCelda.Value := date
	estadoContactoCelda.Value := state
	observacionesCelda.Value := observations

	return True
}

;----------------------------------------
; Atajos
;----------------------------------------
;// Mostrar u ocultar ventanas generadas por el programa.
!F5::
If (WinExist("ahk_id " ieHWND) and !ie.Visible) {
	ie.Visible := True
	WinActivate
	Send, #{Left}
}
Else
	ie.Visible := False

If (WinExist("ahk_id " formHWND) and form.Visible) {
	form.Visible := True
	WinActivate
	Send, #{Right}
}
Else
	form.Visible := False

;If (WinExist("ahk_id " excelHWND) and !excel.Visible) {
;	excel.Visible = True
;	WinMaximize
;}
;Else
;	excel.Visible = False
Return

!F11::Pause

!F12::ExitApp