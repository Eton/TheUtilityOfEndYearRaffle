var X = XLSX;
var emploee, award;

// �w�g���������u�s��
var hasaward = [];

// ��l��
function initial() {

    // initial memory data
    emploee = {};
    award = {};
    hasaward = [];

    // initial html view
    $('#award').empty();
    $('#emploee').empty();
}

// �ɮש즲
var dropaward = document.getElementById('dropaward');
if (dropaward.addEventListener) {
    dropaward.addEventListener('dragenter', handleDragover, false);
    dropaward.addEventListener('dragover', handleDragover, false);
    dropaward.addEventListener('drop', handleDrop, false);
}

// ������s
var gobtn = document.getElementById('go');
if (gobtn.addEventListener) {
    gobtn.addEventListener('click', handleGo, false);
}

var resetbtn = document.getElementById('reset');
if (resetbtn.addEventListener) {
    resetbtn.addEventListener('click', hanleReset, false);
}

var raffleagain = document.getElementById('raffleagain');
if (raffleagain.addEventListener) {
    raffleagain.addEventListener('click', handleRaffleAgain, false);
}

var exportBtn = document.getElementById('export');
if (exportBtn.addEventListener) {
    exportBtn.addEventListener('click', handleExport, false);
}

function handleExport(e) {
    if (typeof award == 'undefined') {
        alert('�|���פJ�������');
        return;
    }

    for (var i = 0 ; i < award.length ; ++i) {
        if (award[i].isAlreadyRaffled === false) {
            alert('�|�����������');
            return;
        }
    }

    export_table_to_excel('emploee');
}

// ��K���եΡA�@�����Ʃ��
function handleRaffleAgain(e) {
    hasaward = [];

    for (var i = 0; i < emploee.length; ++i) {
        emploee[i].���� = "";
        $('#emploee' + emploee[i].�s��).removeClass('hasaward');
        $('#emploee' + (i + 1) + ' .award').html('');
    }

    for (var i = 0; i < award.length; ++i) {
        award[i].isAlreadyRaffled = false;
    }

    handleGo(e);
}

function hanleReset(e) {
    initial();
}

// ������s
function handleGo(e) {

    // �P�_�O�_�����u & �������
    if (typeof emploee == 'undefined' || typeof award == 'undefined' || emploee === null || award === null) {
        alert('�Х��פJ���u & �������');
        return;
    }

    // �ثe�n�⪺�����O
    var awardnum = $('#awardselect').val();

    //var awardObject = $.grep(award, function (e) {
    //    return e.���� == awardname;
    //})[0];
    var awardObject = award[awardnum - 1];

    // �p�G�Ӽ����w�g��L�F
    if (awardObject.isAlreadyRaffled === true) {
        alert(awardObject.���� + " �w�g��L�o!");
        return;
    }

    //for (var i = 0; i < award.length; i++) {
    for (var j = 0; j < awardObject.�ƶq; j++) {
        var num = getRandomNum(1, emploee.length, hasaward);
        markEmploeeHasAward(num, awardObject.����);
        hasaward.push(num);
    }

    markAwardHasRaffled(awardObject.num);
    //}
}

// ���o�üƸ��X
function getRandomNum(min, max, exclude) {
    var num = Math.floor(Math.random() * (max - min + 1)) + min;
    return exclude.indexOf(num) === -1 ? num : getRandomNum(min, max, exclude);
}

// �Хܭ��u�w�g����
function markEmploeeHasAward(num, award) {

    // �P�_�ӭ��u�S���o��
    //if (emploee[num - 1].����)

    emploee[num - 1].���� = award;

    // add has award class
    $('#emploee' + num).addClass('hasaward');

    // add award to emploee table row
    $('#emploee' + num + ' .award').html(award);
}

// �ХܸӼ����w�g��L�F
function markAwardHasRaffled(num) {
    award[num - 1].isAlreadyRaffled = true;

    $('#award' + num + ' .isAlreadyRaffled').html('�w�⧹');
}

function handleDrop(e) {
    e.stopPropagation();
    e.preventDefault();
    var files = e.dataTransfer.files;
    var f = files[0];
    {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function (e) {
            if (typeof console !== 'undefined') console.log("onload", new Date());
            var data = e.target.result;

            var arr = fixdata(data);
            wb = X.read(btoa(arr), { type: 'base64' });

            process_wb(wb);

            setAwardSelect(award);
        }
    };
    reader.readAsArrayBuffer(f);
}

function handleDragover(e) {
    e.stopPropagation();
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
}

function fixdata(data) {
    var o = "", l = 0, w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
}

// �N���u & ������T��ܦb�����W
function process_wb(wb) {
    // �N wordbook �ন json
    var data = to_json(wb);

    // �N���� & ���u �q�b table �W

    // sheet1 ����
    $('#award').empty();
    award = data.����;
    for (var i = 0; i < award.length; ++i) {

        award[i].num = i + 1;
        award[i].isAlreadyRaffled = false;

        tr = $('<tr id="award' + award[i].num + '"/>');
        tr.append("<td>" + award[i].���� + "</td>");
        tr.append("<td>" + award[i].�ƶq + "</td>");
        tr.append("<td class='isAlreadyRaffled'></td>");
        $('#award').append(tr);
    }

    // sheet2 ���u
    $('#emploee').empty();
    emploee = data.���u;
    for (var i = 0; i < emploee.length; ++i) {

        // ���u�s�W�����ݩ�
        emploee[i].�s�� = i + 1;
        emploee[i].���� = "";

        tr = $('<tr id="emploee' + emploee[i].�s�� + '"/>');
        tr.append("<td>" + emploee[i].�s�� + "</td>");
        tr.append("<td>" + emploee[i].���� + "</td>");
        tr.append("<td>" + emploee[i].���u + "</td>");
        tr.append("<td class='award'></td>");
        $('#emploee').append(tr);
    }
}

function setAwardSelect(awards) {
    for (var i = 0; i < awards.length; ++i) {
        $('#awardselect').append($("<option></option>").val(awards[i].num).html(awards[i].����));
    }
}

// �N workbook ��X json
function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function (sheetName) {
        var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if (roa.length > 0) {
            result[sheetName] = roa;
        }
    });
    return result;
}