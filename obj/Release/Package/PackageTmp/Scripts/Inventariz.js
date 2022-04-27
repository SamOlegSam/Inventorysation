// Нажатие вкладки Справка в модальном окне //

function Reference(ID) {
    $.ajax({
        url: "/Home/Reference/",
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}
//----------------------------------------------------//
//-------Добавление члена комиссии--------------//
function AddKomissiya() {
    $.ajax({
        url: "/Home/AddKomissiya/",
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}

//-------Сохранение добавления члена комиссии----------//
function KomissiyaSave() {

    var isValid = true;
    if ($('#komis').val() == "") {
        $('#komis').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#komis').css('border-color', 'lightgrey');
    }

    if ($('#Doljnost').val() == "") {
        $('#Doljnost').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Doljnost').css('border-color', 'lightgrey');
    }

    if ($('#FIO').val() == "") {
        $('#FIO').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#FIO').css('border-color', 'lightgrey');
    }


    if (isValid == false) {
        return false;
    }

    var data = {
        //'ID': ID,
        'location': $('#komis').val(),
        'Nazn': $('#Nazn').val(),
        'Doljnost': $('#Doljnost').val(),
        'FIO': $('#FIO').val()

    };

    $.ajax({
        url: "/Home/KomissiyaSave",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//--------------------------------//
// Удаление члена//

function DeletePodpis(ID) {
    $.ajax({
        url: "/Home/DeletePodpis/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalDeleteContent').html(result);
            $('#ServicesModalDelete').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}
//-----------------------------//
//Подтверждение удаления члена //
function DeletePodpisOK(ID) {


    $.ajax({
        url: "/Home/DeletePodpisOK/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalDeleteContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}
//-----------------------------//
// Редактирование члена комиссии //

function KomissiyaEdit(ID) {
    $.ajax({
        url: "/Home/KomissiyaEdit/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}
//-------------------------//
// Сохранение редактирования члена комиссии //

function KomissiyaEditSave() {

    var isValid = true;

    if ($('#Doljnost').val() == "") {
        $('#Doljnost').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Doljnost').css('border-color', 'lightgrey');
    }

    if ($('#FIO').val() == "") {
        $('#FIO').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#FIO').css('border-color', 'lightgrey');
    }


    if (isValid == false) {
        return false;
    }

    var data = {

        'ID': $('#ID').val(),
        'Nazn': $(Nazn).val(),
        'Doljnost': $('#Doljnost').val(),
        'FIO': $('#FIO').val(),

    };

    $.ajax({
        url: "/Home/KomissiyaEditSave",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}
//-----------------Получение градуировочной таблицы в зависимости от резервуара-------------------------

function GetGradTable() {

    var isValid = true;

    if ($('#rezer').val() == "0") {
        $('#rezer').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#rezer').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {

        'rezer': $('#rezer').val(),
    };

    $.ajax({
        url: "/Home/GetGradTable",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#GradTab').html(result);

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//------------------------------------------------------------------------------
// Удаление строки в градуировочной таблице//

function DeleteGrad(ID) {
    $.ajax({
        url: "/Home/DeleteGrad/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalDeleteContent').html(result);
            $('#ServicesModalDelete').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}
//-----------------------------//
//Подтверждение удаления строки в градуировочной таблице //
function DeleteGradOK(ID) {


    $.ajax({
        url: "/Home/DeleteGradOK/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalDeleteContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}
//-----------------------------//
// Редактирование градуировочной таблицы //

function GradEdit(ID) {
    $.ajax({
        url: "/Home/GradEdit/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}
//-------------------------//
// Сохранение редактирования градуировочной таблицы //

function GradEditSave() {

    var isValid = true;

    if ($('#urov').val() == "") {
        $('#urov').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#urov').css('border-color', 'lightgrey');
    }

    if ($('#V').val() == "") {
        $('#V').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#V').css('border-color', 'lightgrey');
    }


    if (isValid == false) {
        return false;
    }

    var data = {

        'ID': $('#ID').val(),
        'urov': $('#urov').val(),
        'V': $('#V').val(),

    };

    $.ajax({
        url: "/Home/GradEditSave",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}

//-------Добавление строки в градуировочную таблицу--------------//
function AddGradTable() {

    var isValid = true;
    if ($('#rezer').val() == "0") {
        $('#rezer').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#rezer').css('border-color', 'lightgrey');
    }       

    if (isValid == false) {
        return false;
    }
    $.ajax({
        url: "/Home/AddGradTable/",
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}

//-------Сохранение добавления строки в градуировочную таблицу----------//
function GradTableSave() {

    var isValid = true;
    if ($('#urov').val() == "") {
        $('#urov').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#urov').css('border-color', 'lightgrey');
    }

    if ($('#V').val() == "") {
        $('#V').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#V').css('border-color', 'lightgrey');
    }
       
    if (isValid == false) {
        return false;
    }

    var data = {
        //'ID': ID,
        'IDRez': $('#rezer').val(),
        'urov': $('#urov').val(),
        'V': $('#V').val(),
        
    };

    $.ajax({
        url: "/Home/GradTableSave",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//--------------------------------//
//-----------------Получение химанализа в зависимости от резервуара-------------------------

function GetHim() {

    var isValid = true;

    if ($('#filial').val() == "") {
        $('#filial').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#filial').css('border-color', 'lightgrey');
    }

    if ($('#rezer').val() == "0") {
        $('#rezer').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#rezer').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {
        'filial': $('#filial').val(),
        'rezer': $('#rezer').val(),
        
    };

    $.ajax({
        url: "/Home/GetHim",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            
            $('#res').html(result)
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//------------------------------------------------------------------------------
//-----------------Передача параметров для калькулятора рассчет разности объемов-------------------------

function GetRazn() {

    var isValid = true;

    if ($('#filial1').val() == "") {
        $('#filial1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#filial1').css('border-color', 'lightgrey');
    }

    if ($('#rezer1').val() == "0") {
        $('#rezer1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#rezer1').css('border-color', 'lightgrey');
    }

    if ($('#Unach').val() == "") {
        $('#Unach').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Unach').css('border-color', 'lightgrey');
    }

    if ($('#Ukon').val() == "") {
        $('#Ukon').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Ukon').css('border-color', 'lightgrey');
    }

    if ($('#Pnach').val() == "") {
        $('#Pnach').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Pnach').css('border-color', 'lightgrey');
    }

    if ($('#Pkon').val() == "") {
        $('#Pkon').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Pkon').css('border-color', 'lightgrey');
    }

    if ($('#Tnach').val() == "") {
        $('#Tnach').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Tnach').css('border-color', 'lightgrey');
    }

    if ($('#Tkon').val() == "") {
        $('#Tkon').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Tkon').css('border-color', 'lightgrey');
    }

    if ($('#Bal').val() == "") {
        $('#Bal').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Bal').css('border-color', 'lightgrey');
    }

    if ($('#Kst').val() == "") {
        $('#Kst').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Kst').css('border-color', 'lightgrey');
    }

    if ($('#Ksr').val() == "") {
        $('#Ksr').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Ksr').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {
        'filial1': $('#filial1').val(),
        'rezer1': $('#rezer1').val(),
        'Unach': $('#Unach').val(),
        'Ukon': $('#Ukon').val(),
        'UnachH2O': $('#UnachH2O').val(),
        'UkonH2O': $('#UkonH2O').val(),
        'Pnach': $('#Pnach').val(),
        'Pkon': $('#Pkon').val(),
        'Tnach': $('#Tnach').val(),
        'Tkon': $('#Tkon').val(),
        'Bal': $('#Bal').val(),
        'Kst': $('#Kst').val(),
        'Ksr': $('#Ksr').val(),

    };

    $.ajax({
        url: "/Home/GetRazn",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            
            $('#razn').html(result)
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//------------------------------------------------------------------------------

//-----------------Передача параметров для калькулятора рассчет по массе-------------------------

function GetRazn1() {

    var isValid = true;

    if ($('#filial2').val() == "") {
        $('#filial2').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#filial2').css('border-color', 'lightgrey');
    }

    if ($('#rezer2').val() == "0") {
        $('#rezer2').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#rezer2').css('border-color', 'lightgrey');
    }

    if ($('#Unach1').val() == "") {
        $('#Unach1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Unach1').css('border-color', 'lightgrey');
    }

    if ($('#RM1').val() == "") {
        $('#RM1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#RM1').css('border-color', 'lightgrey');
    }

    if ($('#Pnach1').val() == "") {
        $('#Pnach1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Pnach1').css('border-color', 'lightgrey');
    }

    if ($('#Pkon1').val() == "") {
        $('#Pkon1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Pkon1').css('border-color', 'lightgrey');
    }

    if ($('#Tnach1').val() == "") {
        $('#Tnach1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Tnach1').css('border-color', 'lightgrey');
    }

    if ($('#Tkon1').val() == "") {
        $('#Tkon1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#Tkon1').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {
        'filial2': $('#filial2').val(),
        'rezer2': $('#rezer2').val(),
        'Unach1': $('#Unach1').val(),
        'UnachH2O1': $('#UnachH2O1').val(),
        'UkonH2O1': $('#UkonH2O1').val(),
        'RM1': $('#RM1').val(),
        'Pnach1': $('#Pnach1').val(),
        'Pkon1': $('#Pkon1').val(),
        'Tnach1': $('#Tnach1').val(),
        'Tkon1': $('#Tkon1').val(),
        'Bal1': $('#Bal1').val(),
        'Kst1': $('#Kst1').val(),
        'Ksr1': $('#Ksr1').val(),
    };

    $.ajax({
        url: "/Home/GetRazn1",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {

            $('#razn1').html(result)
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });

}
//------------------------------------------------------------------------------
//--------------Получение таблицы инвентаризации в зависимости от даты и филиала----------------------------------------------

function GetTableInv() {

    var isValid = true;

    if ($('#filial1').val() == "0") {
        $('#filial1').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#filial1').css('border-color', 'lightgrey');
    }

    if ($('#datinv').val() == "0") {
        $('#datinv').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#datinv').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {

        'filial1': $('#filial1').val(),
        'datinv': $('#datinv').val(),
    };

    $.ajax({
        url: "/Home/GetTableInv",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#Tableinv').html(result);

        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
 }
//-----Формирование отчета в EXCEL-----------------------------------------------
function Report() {
    var stringhref = "/Home/Export?";

    stringhref += "filial1=" + $('#filial1').val() + "&" + "datinv=" + $('#datinv').val();
    window.location = stringhref;
}


//-----------------------------//
// Редактирование инвентаризации //

function InventEdit(ID) {
    $.ajax({
        url: "/Home/InventEdit/" + ID,
        type: "GET",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
            $('#ServicesModal').modal('show');
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }

    })
}
//-------------------------//
// Сохранение редактирования инвентаризации //

function InventEditSave() {

    var isValid = true;

    if ($('#InvRezerEdit').val() == "") {
        $('#InvRezerEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvRezerEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvUrovEdit').val() == "") {
        $('#InvUrovEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvUrovEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvUrovH2OEdit').val() == "") {
        $('#InvUrovH2OEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvUrovH2OEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvVEdit').val() == "") {
        $('#InvVEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvVEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvVH2OEdit').val() == "") {
        $('#InvVH2OEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvVH2OEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvVNeftEdit').val() == "") {
        $('#InvVNeftEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvVNeftEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvTempEdit').val() == "") {
        $('#InvTempEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvTempEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvPEdit').val() == "") {
        $('#InvPEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvPEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvMassaBruttoEdit').val() == "") {
        $('#InvMassaBruttoEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvMassaBruttoEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvH2OEdit').val() == "") {
        $('#InvH2OEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvH2OEdit').css('border-color', 'lightgrey');
    }


    if ($('#InvSaltEdit').val() == "") {
        $('#InvSaltEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvSaltEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvMehEdit').val() == "") {
        $('#InvMehEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvMehEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvBalProcEdit').val() == "") {
        $('#InvBalProcEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvBalProcEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvBalTonnEdit').val() == "") {
        $('#InvBalTonnEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvBalTonnEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvMassaNettoEdit').val() == "") {
        $('#InvMassaNettoEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvMassaNettoEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvHMimEdit').val() == "") {
        $('#InvHMimEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvHMimEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvVMinEdit').val() == "") {
        $('#InvVMinEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvVMinEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvMNettoMinEdit').val() == "") {
        $('#InvMNettoMinEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvMNettoMinEdit').css('border-color', 'lightgrey');
    }

    if ($('#InvVTehEdit').val() == "") {
        $('#InvVTehEdit').css('border-color', 'Red');
        isValid = false;
    }
    else {
        $('#InvVTehEdit').css('border-color', 'lightgrey');
    }

    if (isValid == false) {
        return false;
    }

    var data = {

        'ID': $('#InvIDEdit').val(),
        'InvRezer': $('#InvRezerEdit').val(),
        'InvUrov': $('#InvUrovEdit').val(),
        'InvUrovH2O': $('#InvUrovH2OEdit').val(),
        'InvV': $('#InvVEdit').val(),
        'InvVH2O': $('#InvVH2OEdit').val(),
        'InvVNeft': $('#InvVNeftEdit').val(),
        'InvTemp': $('#InvTempEdit').val(),
        'InvP': $('#InvPEdit').val(),
        'InvMassaBrutto': $('#InvMassaBruttoEdit').val(),
        'InvH2O': $('#InvH2OEdit').val(),
        'InvSalt': $('#InvSaltEdit').val(),
        'InvMeh': $('#InvMehEdit').val(),
        'InvBalProc': $('#InvBalProcEdit').val(),
        'InvBalTonn': $('#InvBalTonnEdit').val(),
        'InvMassaNetto': $('#InvMassaNettoEdit').val(),
        'InvHMim': $('#InvHMimEdit').val(),
        'InvVMin': $('#InvVMinEdit').val(),
        'InvMNettoMin': $('#InvMNettoMinEdit').val(),
        'InvVTeh': $('#InvVTehEdit').val(),
    };

    $.ajax({
        url: "/Home/InventEditSave",
        type: "POST",
        contentType: "application/json;charset=UTF-8",
        data: JSON.stringify(data),
        dataType: "html",
        success: function (result) {
            $('#ServicesModalContent').html(result);
        },
        error: function (errormessage) {
            alert(errormessage.responseText);
        }
    });
}