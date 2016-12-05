var X = XLSX;
var emploee, award;

// 已經中獎的員工編號
//var hasaward = [];

//var raffleTimeInterval = 1500;

var slotmachine;

// 抽獎次數
var raffleCount = 0;

// 目前抽獎的獎項名稱
//var currentAwardName;

// 目前抽獎的獎項編號
//var currentAwardNum;

// 目前抽獎的獎項資訊
var currentAward;

// 避免重啟
function closeIt() {
    return "確定要關閉此頁面嗎";
}

window.onbeforeunload = closeIt;

// 初始化
function initial() {

    // initial memory data
    emploee = [];
    award = [];
    raffleCount = 0;
    //hasaward = [];

    // initial html view
    $('#award tbody').empty();
    $('#emploee tbody').empty();
    $('#awardselect').empty();

    // 顯示可以丟置檔案的框框
    $('#dropaward').show();

    // 重置 slot machine 的隨機數字
    slotmachine[0].slotmachine.InitialRadomNum();
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

//var raffleagain = document.getElementById('raffleagain');
//if (raffleagain.addEventListener) {
//    raffleagain.addEventListener('click', handleRaffleAgain, false);
//}

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

    export_table_to_excel('emploee', "raffle.xlsx");
}

// 方便測試用，一直重複抽獎
//function handleRaffleAgain(e) {
//    //hasaward = [];

//    for (var i = 0; i < emploee.length; ++i) {
//        emploee[i].獎項 = "";
//        $('#emploee' + emploee[i].編號).css("background-color", "white").css("color", "black");
//        $('#emploee' + (i + 1) + ' .award').html('');
//    }

//    for (var i = 0; i < award.length; ++i) {
//        award[i].isAlreadyRaffled = false;
//    }

//    handleGo(e);
//}

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
    //var awardObject = award[awardnum - 1];
    currentAward = award[awardnum - 1];

    // 如果該獎項已經抽過了
    if (currentAward.isAlreadyRaffled === true) {
        alert(currentAward.獎項 + " 已經抽過囉!");
        return;
    }

    raffleCount = currentAward.數量;
    //currentAwardName = awardObject.獎項;
    //currentAwardNum = awardObject.num;
    //currentAward = awardObject;

    //$('#slotfunction').css('opacity', 1);
    //$('#slotfunction').show();
    //openSlotMachine();

    // 開始抽獎
    raffle();
}

//function openSlotMachine() {
//    $('#slotmachine').css('opacity', 1);
//    $('#slotmachinebackground').show();
//    $('#slotbackground').show();
//}

//function closeSlotMachine() {
//    $('#slotmachine').css('opacity', 0);
//    $('#slotmachinebackground').hide();
//    $('#slotbackground').hide();
//}

// 抽獎
function raffle() {
    if (raffleCount-- <= 0) {
        //$('#slotfunction').css('opacity',0);
        //$('#slotfunction').hide();
        //closeSlotMachine();
        markAwardHasRaffled(currentAward.編號);

        export_table_to_excel('emploee', currentAward.獎項 + ".xlsx");

        return;
    }

    slotmachine[0].slotmachine.playSlots(currentAward.表演時間);
}

function raffleEnd(raffleNum) {
    //alert(raffleNum);
    playCongratulationsSoundEffects();

    // 移動scollbar
    moveEmploeeScrollbar(raffleNum);
}

function moveEmploeeScrollbar(raffleNum) {

    // 取得 table 高度
    var tableHeight = $('#emploeeDiv').height();

    // 取得一個 row 的高度
    var unitHeight = $('#emploee tbody tr').eq(0).height();

    // 計算目標高度
    var targetTop = (raffleNum - 1) * unitHeight;

    // 超過 table 的一半高度才移動
    targetTop = (targetTop > (tableHeight / 2)) ? targetTop - (tableHeight / 2) : 0;

    // 目前高度
    var currentTop = $("#emploeeDiv").scrollTop();

    // 依照移動高度差來設定時間
    var animationTime = Math.abs(currentTop - targetTop) / unitHeight * 40;

    $("#emploeeDiv").animate({ scrollTop: targetTop }, animationTime, 'easeOutSine', function () {
        markEmploeeHasAward(raffleNum, currentAward);
        raffle();
    });
}

function sleep(milliseconds) {
    var start = new Date().getTime();
    for (var i = 0; i < 1e7; i++) {
        if ((new Date().getTime() - start) > milliseconds) {
            break;
        }
    }
}

// 取得亂數號碼
function getRandomNum(min, max, exclude) {
    var num = Math.floor(Math.random() * (max - min + 1)) + min;
    return exclude.indexOf(num) === -1 ? num : getRandomNum(min, max, exclude);
}

// 標示員工已經中獎
function markEmploeeHasAward(raffleNum, award) {

    // 判斷該員工沒有得獎
    //if (emploee[num - 1].獎項)

    emploee[raffleNum - 1].獎項 = award.編號;

    // add has award class
    //$('#emploee' + num).addClass('hasaward');
    $('#emploee' + raffleNum).css("background-color", award.背景顏色).css("color", award.文字顏色);

    // add award to emploee table row
    $('#emploee' + raffleNum + ' .award').html(award.獎項);
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

            $('#dropaward').hide();

            process_wb(wb);

            setAwardColor();

            setAwardSelect(award);

            initialSlotMachine();
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

    // 按照獎項編號排序
    award = award.sort(function (a, b) {
        if (a.編號 > b.編號) return 1;
        if (a.編號 <= b.編號) return -1;
        return 0;
    })

    for (var i = 0; i < award.length; ++i) {

        //award[i].編號 = i + 1;
        award[i].isAlreadyRaffled = false;

        tr = $('<tr id="award' + award[i].編號 + '"/>');
        tr.append("<td>" + award[i].編號 + "</td>");
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
}

function initialSlotMachine() {

    if (checkIsDefined(slotmachine))
        return;

    $('#slotmachine').show();

    slotmachine = $('.slot').jSlots({
        onEnd: raffleEnd,
        min: 1,
        max: emploee.length,
    });

    //$('#slotmachine').css('opacity', 1);
}


// 設定獎項 html 背景顏色 & 文字顏色
function setAwardColor() {
    for (var i = 1; i < award.length + 1; ++i) {
        $('#award' + i).css("background-color", award[i - 1].背景顏色).css("color", award[i - 1].文字顏色);
    }
}

function setAwardSelect(awards) {
    $('#awardselect').empty();

    for (var i = awards.length - 1; i >= 0; --i) {
        $('#awardselect').append($("<option></option>").val(awards[i].編號).html("(" + (awards[i].編號) + ") " + awards[i].獎項));
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

    //var num = getRandomNum(1, 2, []);
    if (checkIsDefined($('#Congratulations')[0]) === true)
        $('#Congratulations')[0].play();
}

function checkIsDefined(object) {
    if (typeof object !== 'undefined')
        return true;
    else
        return false;
}

function log(message) {
    //var txtFile = "../test.txt";
    //var file = new File([""], txtFile);

    //file.open("w"); // open file with write access    
    //file.writeln(message);
    //file.writeln("-------------------------------------");
    //file.close();

    var fso = new CreateObject("Scripting.FileSystemObject");
    var fh = fso.CreateTextFile("../test.txt", 8, true, 0);
    fh.WriteLine(message);
    fh.WriteLine("-------------------------------------");
    fh.Close();
}

//initial();

//log("Begin");
