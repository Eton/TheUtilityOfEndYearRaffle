var X = XLSX;
var emploee, award;

// 已經中獎的員工編號
var hasaward = [];

var raffleTimeInterval = 1500;

var slotmachine;

// 抽獎次數
var raffleCount = 0;

// 目前抽獎的獎項名稱
var currentAwardName;

// 目前抽獎的獎項編號
var currentAwardNum;

// 初始化
function initial() {

    // initial memory data
    emploee = [];
    award = [];
    hasaward = [];

    // initial html view
    $('#award tbody').empty();
    $('#emploee tbody').empty();
    $('#awardselect').empty();    
}

// 檔案拖曳
var dropaward = document.getElementById('dropaward');
if (dropaward.addEventListener) {
    dropaward.addEventListener('dragenter', handleDragover, false);
    dropaward.addEventListener('dragover', handleDragover, false);
    dropaward.addEventListener('drop', handleDrop, false);
}

// 抽獎按鈕
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
        alert('尚未匯入獎項資料');
        return;
    }

    for (var i = 0 ; i < award.length ; ++i) {
        if (award[i].isAlreadyRaffled === false) {
            alert('尚有獎項未抽獎');
            return;
        }
    }

    export_table_to_excel('emploee');
}

// 方便測試用，一直重複抽獎
function handleRaffleAgain(e) {
    hasaward = [];

    for (var i = 0; i < emploee.length; ++i) {
        emploee[i].獎項 = "";
        $('#emploee' + emploee[i].編號).removeClass('hasaward');
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

// 抽獎按鈕
function handleGo(e) {

    // 判斷是否有員工 & 獎項資料
    if (typeof emploee == 'undefined' || typeof award == 'undefined' || emploee === null || award === null || emploee.length <= 0 || award.length <= 0) {
        alert('請先匯入員工 & 獎項資料');
        return;
    }

    // 目前要抽的獎項是
    var awardnum = $('#awardselect').val();

    // 獎項資料
    var awardObject = award[awardnum - 1];

    // 如果該獎項已經抽過了
    if (awardObject.isAlreadyRaffled === true) {
        alert(awardObject.獎項 + " 已經抽過囉!");
        return;
    }

    raffleCount = awardObject.數量;
    currentAwardName = awardObject.獎項;
    currentAwardNum = awardObject.num;

    // 開始抽獎
    raffle();
}

// 抽獎
function raffle() {
    if (raffleCount-- <= 0)
    {
        markAwardHasRaffled(currentAwardNum);
        return;
    }
        
    slotmachine[0].slotmachine.playSlots();
}

function raffleEnd(raffleNum) {
    alert(raffleNum);
    playCongratulationsSoundEffects();
    markEmploeeHasAward(raffleNum, currentAwardName);
    raffle();
}

// 取得亂數號碼
function getRandomNum(min, max, exclude) {
    var num = Math.floor(Math.random() * (max - min + 1)) + min;
    return exclude.indexOf(num) === -1 ? num : getRandomNum(min, max, exclude);
}

// 標示員工已經中獎
function markEmploeeHasAward(num, award) {

    // 判斷該員工沒有得獎
    //if (emploee[num - 1].獎項)

    emploee[num - 1].獎項 = award;

    // add has award class
    $('#emploee' + num).addClass('hasaward');

    // add award to emploee table row
    $('#emploee' + num + ' .award').html(award);
}

// 標示該獎項已經抽過了
function markAwardHasRaffled(num) {
    award[num - 1].isAlreadyRaffled = true;

    $('#award' + num + ' .isAlreadyRaffled').html('已抽完');
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

// 將員工 & 獎項資訊顯示在頁面上
function process_wb(wb) {
    // 將 wordbook 轉成 json
    var data = to_json(wb);

    // 將獎項 & 員工 秀在 table 上

    // sheet1 獎項
    $('#award tbody').empty();

    award = data.獎項;
    for (var i = 0; i < award.length; ++i) {

        award[i].num = i + 1;
        award[i].isAlreadyRaffled = false;

        tr = $('<tr id="award' + award[i].num + '"/>');
        tr.append("<td>" + award[i].num + "</td>");
        tr.append("<td>" + award[i].獎項 + "</td>");
        tr.append("<td>" + award[i].數量 + "</td>");
        tr.append("<td class='isAlreadyRaffled'></td>");
        $('#award tbody').append(tr);
    }

    // sheet2 員工
    $('#emploee tbody').empty();

    emploee = data.員工;
    for (var i = 0; i < emploee.length; ++i) {

        // 員工新增獎項屬性
        emploee[i].編號 = i + 1;
        emploee[i].獎項 = "";

        tr = $('<tr id="emploee' + emploee[i].編號 + '"/>');
        tr.append("<td>" + emploee[i].編號 + "</td>");
        tr.append("<td>" + emploee[i].部門 + "</td>");
        tr.append("<td>" + emploee[i].員工 + "</td>");
        tr.append("<td class='award'></td>");
        $('#emploee tbody').append(tr);
    }

    initialSlotMachine();
}

function initialSlotMachine() {
    
    $('#slotmachine').show();

    slotmachine = $('.slot').jSlots({
        onEnd: raffleEnd,
        min: 1,
        max: emploee.length,
    });    
}

function setAwardSelect(awards) {
    $('#awardselect').empty();

    for (var i = 0; i < awards.length; ++i) {
        $('#awardselect').append($("<option></option>").val(awards[i].num).html(awards[i].獎項));
    }
}

// 將 workbook 輸出 json
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

// 播放恭喜中獎音效
function playCongratulationsSoundEffects() {

    var num = getRandomNum(1, 2, []);
    $('#Congratulations' + num)[0].play();
}

function test() {
    $('#content').load('raffle.html');
}

initial();