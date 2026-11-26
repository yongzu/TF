/*
====================================================================
  03_LengthTest_2up.jsx — 행장 테스트 (좌/우 2단 레이아웃)
  Typography as Foundation 사전 과제

  한 페이지에 왼쪽/오른쪽 두 개의 행장(글줄길이) 값을 나란히 배치합니다.
  글자 수(20~36자)를 실제 프레임 폭(mm)으로 환산해 body_L/body_R의
  폭을 직접 조정합니다. (왼쪽 프레임의 왼쪽 끝, 오른쪽 프레임의 오른쪽
  끝은 템플릿 그대로 두고, 반대쪽 끝만 움직여서 폭을 바꿉니다.)

  【사전 준비 - 최초 1회만】 (01, 02번 스크립트와 동일한 템플릿 라벨 사용)
    - 왼쪽 본문 텍스트 프레임        → body_L
    - 오른쪽 본문 텍스트 프레임      → body_R
    - 왼쪽 하단 정보표(크기/행송/행장) → infobar_L
    - 오른쪽 하단 정보표             → infobar_R
    - 가운데 페이지 번호 박스        → pagenum_center

  【사용법】
  1. 01, 02번 스크립트에서 뽑은 활자 크기 / 행송 값을 아래에 입력한다.
  2. [창 > 유틸리티 > 스크립트] 패널에서 이 파일을 더블클릭 실행한다.

  【참고】
  글자 수 → mm 환산은 "가" 한 글자의 실제 폭을 임시로 측정해서 계산한
  근사치입니다. 서체에 따라 오차가 있을 수 있으니, 생성된 페이지를 보고
  필요하면 CHAR_WIDTH_FACTOR를 0.9~1.1 사이에서 조정하세요.
====================================================================
*/

#target indesign
#targetengine "session"

function mm2pt(mm) { return mm * 2.834645669; }
function pt2mm(pt) { return pt / 2.834645669; }
function roundTo(val, digits) { var f = Math.pow(10, digits); return Math.round(val * f) / f; }

// ===================== 사용자 설정 =====================
var FIXED_SIZE = 10;             // ← 01번 스크립트에서 추출한 활자 크기(pt)
var FIXED_LEADING_PERCENT = 170; // ← 02번 스크립트에서 추출한 행송(%)
var START_CHARS = 20;             // 행장 시작값 (글자 수)
var END_CHARS = 36;               // 행장 끝값 (글자 수)
var STEP = 1;                      // 증가 단위 (글자 수)
var CHAR_WIDTH_FACTOR = 1.0;       // 글자 폭 보정 계수
var CLEAR_EXTRA_PAGES_BEFORE_RUN = true;

var LABELS = {
    bodyL: "body_L",
    bodyR: "body_R",
    infobarL: "infobar_L",
    infobarR: "infobar_R",
    pagenum: "pagenum_center"
};
// ========================================================

function main() {
    if (app.documents.length === 0) {
        alert("먼저 콤핑 템플릿 문서를 열어주세요.");
        return;
    }
    var doc = app.activeDocument;
    var templatePage = doc.pages.item(0);

    var required = [LABELS.bodyL, LABELS.bodyR, LABELS.infobarL, LABELS.infobarR];
    for (var r = 0; r < required.length; r++) {
        if (!findByLabelInPage(templatePage, required[r])) {
            alert("템플릿(1페이지)에서 라벨 '" + required[r] + "' 이 붙은 객체를 찾을 수 없습니다.\n" +
                "창 > 유틸리티 > 스크립트 레이블 패널에서 라벨을 먼저 붙여주세요.");
            return;
        }
    }

    var sourceBodyL = findByLabelInPage(templatePage, LABELS.bodyL);
    var corpusText = sourceBodyL.contents;
    var appliedFont = sourceBodyL.texts[0].appliedFont;
    var leadingPt = roundTo(FIXED_SIZE * (FIXED_LEADING_PERCENT / 100), 2);

    var charsList = [];
    for (var c = START_CHARS; c <= END_CHARS; c += STEP) {
        charsList.push(c);
    }
    var totalPages = Math.ceil(charsList.length / 2);

    app.doScript(function () {
        if (CLEAR_EXTRA_PAGES_BEFORE_RUN) clearExtraPages(doc);

        // "가" 한 글자의 실제 폭(pt) 측정 (1페이지=템플릿 위에서 임시로 측정 후 제거)
        var charWidthPt = measureCharWidth(templatePage, appliedFont, FIXED_SIZE) * CHAR_WIDTH_FACTOR;

        var pageCount = 0;
        for (var i = 0; i < charsList.length; i += 2) {
            var leftChars = charsList[i];
            var rightChars = (i + 1 < charsList.length) ? charsList[i + 1] : null;
            pageCount++;

            var page = templatePage.duplicate(LocationOptions.AT_END);

            var bodyL = findByLabelInPage(page, LABELS.bodyL);
            var bodyR = findByLabelInPage(page, LABELS.bodyR);
            var infobarL = findByLabelInPage(page, LABELS.infobarL);
            var infobarR = findByLabelInPage(page, LABELS.infobarR);
            var pagenum = findByLabelInPage(page, LABELS.pagenum);

            // 왼쪽
            bodyL.contents = corpusText;
            applySizeLeading(bodyL, appliedFont, FIXED_SIZE, leadingPt);
            var leftWidthPt = charWidthPt * leftChars;
            resizeFrameWidth(bodyL, leftWidthPt, "right"); // 왼쪽 끝 고정, 오른쪽 끝만 이동
            var leftMM = roundTo(pt2mm(leftWidthPt), 1);
            setInfobar(infobarL, FIXED_SIZE, FIXED_LEADING_PERCENT, leftChars, leftMM);

            // 오른쪽
            if (rightChars !== null) {
                bodyR.visible = true;
                infobarR.visible = true;
                bodyR.contents = corpusText;
                applySizeLeading(bodyR, appliedFont, FIXED_SIZE, leadingPt);
                var rightWidthPt = charWidthPt * rightChars;
                resizeFrameWidth(bodyR, rightWidthPt, "left"); // 오른쪽 끝 고정, 왼쪽 끝만 이동
                var rightMM = roundTo(pt2mm(rightWidthPt), 1);
                setInfobar(infobarR, FIXED_SIZE, FIXED_LEADING_PERCENT, rightChars, rightMM);
            } else {
                bodyR.visible = false;
                infobarR.visible = false;
            }

            if (pagenum) setPageNumber(pagenum, pageCount, totalPages);
        }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "행장 테스트(2단)");

    alert("행장 테스트 페이지 생성 완료. (" + totalPages + "페이지, 값 " + charsList.length + "개)");
}

function applySizeLeading(frame, appliedFont, pointSize, leadingPt) {
    var t = frame.texts[0];
    t.appliedFont = appliedFont;
    t.pointSize = pointSize;
    t.leading = leadingPt;
}

// anchorSide: "right" → 왼쪽(x1) 고정, 오른쪽(x2)만 새 폭에 맞게 이동
//             "left"  → 오른쪽(x2) 고정, 왼쪽(x1)만 새 폭에 맞게 이동
function resizeFrameWidth(frame, newWidthPt, anchorSide) {
    var b = frame.geometricBounds; // [y1, x1, y2, x2]
    if (anchorSide === "right") {
        frame.geometricBounds = [b[0], b[1], b[2], b[1] + newWidthPt];
    } else {
        frame.geometricBounds = [b[0], b[3] - newWidthPt, b[2], b[3]];
    }
}

// 임시 텍스트 프레임을 만들어 "가" 한 글자의 실제 폭(pt)을 측정 후 제거
function measureCharWidth(page, appliedFont, pointSize) {
    var tempFrame = page.textFrames.add();
    tempFrame.contents = "가";
    tempFrame.texts[0].appliedFont = appliedFont;
    tempFrame.texts[0].pointSize = pointSize;
    tempFrame.fit(FitOptions.FRAME_TO_CONTENT);
    var bounds = tempFrame.geometricBounds; // [y1, x1, y2, x2]
    var width = bounds[3] - bounds[1];
    tempFrame.remove();
    return width;
}

var COL_SIZE_VAL = 1;
var COL_LEAD_PT_VAL = 3;
var COL_LEAD_PCT_VAL = 4;
var COL_LEN_CHAR_VAL = 6;
var COL_LEN_MM_VAL = 7;

function setInfobar(frame, pointSize, leadingPercent, charCount, mmValue) {
    var sizeStr = pointSize + "pt";
    var leadPtStr = roundTo(pointSize * (leadingPercent / 100), 2) + "pt";
    var leadPctStr = leadingPercent + "%";
    var charStr = charCount + "자";
    var mmStr = mmValue + "mm";

    if (frame.tables.length > 0) {
        var row = frame.tables[0].rows[0];
        row.cells[COL_SIZE_VAL].texts[0].contents = sizeStr;
        row.cells[COL_LEAD_PT_VAL].texts[0].contents = leadPtStr;
        row.cells[COL_LEAD_PCT_VAL].texts[0].contents = leadPctStr;
        row.cells[COL_LEN_CHAR_VAL].texts[0].contents = charStr;
        row.cells[COL_LEN_MM_VAL].texts[0].contents = mmStr;
    } else {
        frame.contents = "크기\t" + sizeStr + "\t행송\t" + leadPtStr + "\t" + leadPctStr +
            "\t행장\t" + charStr + "\t" + mmStr;
    }
}

function setPageNumber(frame, current, total) {
    var t = frame.texts[0];
    if (t.paragraphs.length >= 1) {
        t.paragraphs.item(0).contents = current + " / " + total;
    } else {
        frame.contents = current + " / " + total;
    }
}

function findByLabelInPage(page, label) {
    var found = null;
    function search(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.label === label) { found = item; return true; }
            if (item.pageItems && item.pageItems.length > 0) {
                if (search(item.pageItems)) return true;
            }
        }
        return false;
    }
    search(page.pageItems);
    return found;
}

function clearExtraPages(doc) {
    for (var i = doc.pages.length - 1; i >= 1; i--) {
        doc.pages[i].remove();
    }
}

main();
