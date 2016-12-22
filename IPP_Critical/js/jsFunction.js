function SearchList(con1, con2) {
    
    //var tb = document.getElementById('<%= txb_lotinput.ClientID %>');
    //var l = document.getElementById('<%= lb_lotSource.ClientID %>');

    var tb = document.getElementById(con1);
    var l = document.getElementById(con2);

    if (tb.value == "") {
        l.selectedIndex = -1;
    }
    else {
        for (var i = 0; i < l.options.length; i++) {
            if (l.options[i].value.toLowerCase().match(tb.value.toLowerCase())) {
                l.selectedIndex = i;
                l.options[i].selected = true;
                return false;
            }
            else {
                l.selectedIndex = -1;
            }
        }
    }
}
// --- 左右移動 ---
function moveSelected(con1, con2) {

    var oSourceSel = document.getElementById(con1);
    var oTargetSel = document.getElementById(con2);
    var arrSelValue = new Array();
    var arrSelText = new Array();
    var arrValueTextRelation = new Array();
    var index = 0;

    for (var i = 0; i < oSourceSel.options.length; i++) {
        if (oSourceSel.options[i].selected) {
            arrSelValue[index] = oSourceSel.options[i].value;
            arrSelText[index] = oSourceSel.options[i].text;
            arrValueTextRelation[arrSelValue[index]] = oSourceSel.options[i];
            index++;
        }
    }

    for (var i = 0; i < arrSelText.length; i++) {
        var oOption = document.createElement("option");
        oOption.text = arrSelText[i];
        oOption.value = arrSelValue[i];
        oTargetSel.add(oOption);
        oSourceSel.removeChild(arrValueTextRelation[arrSelValue[i]]);
    }
}

function moveAll(oSourceSel, oTargetSel) {
    var arrSelValue = new Array();
    var arrSelText = new Array();

    for (var i = 0; i < oSourceSel.options.length; i++) {
        arrSelValue[i] = oSourceSel.options[i].value;
        arrSelText[i] = oSourceSel.options[i].text;
    }

    for (var i = 0; i < arrSelText.length; i++) {
        var oOption = document.createElement("option");
        oOption.text = arrSelText[i];
        oOption.value = arrSelValue[i];
        oTargetSel.add(oOption);
    }

    oSourceSel.innerHTML = "";
}

function deleteSelectItem(oSelect) {
    for (var i = 0; i < oSelect.options.length; i++) {
        if (i >= 0 && i <= oSelect.options.length - 1 && oSelect.options[i].selected) {
            oSelect.options[i] = null;
            i--;
        }
    }
}