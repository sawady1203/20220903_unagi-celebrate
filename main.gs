function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // シート名を取得
  const quizNumber = sheet.getSheetName();
  console.log("シート名（問題）：", quizNumber, typeof quizNumber);

  // 最終行を数値で返す
  const lastRow = sheet.getLastRow();
  console.log("行数：", lastRow, typeof lastRow);

  // 新しい入力データを取得
  const lastData = sheet.getRange(lastRow, 1, 1, 7).getValues()[0];
  console.log("新しいデータ:", lastData);
  // 新しいデータ:
  // [ Sat Aug 13 2022 01:19:46 GMT+0900 (Japan Standard Time),
  // 'たなか',
  // 'たろう',
  // 4,
  // 3,
  // 2,
  // 1 ]

  // クイズのメタ情報から取得
  const quiz_meta = getQuizMeta(ss, quizNumber);
  console.log("クイズメタ：", quiz_meta, typeof quiz_meta);

  // 入力時間の最終行を取得
  const inputTime = lastData[0];
  console.log("入力時間：", inputTime, typeof inputTime);

  //<-- 氏名 -->
  const full_name = lastData.slice(1, 3).join(" ");
  console.log("回答者氏名：", full_name, typeof full_name);

  //<-- 最終回答 -->
  const final_answer = lastData.slice(3, 7).join("");
  console.log("最終回答：", final_answer, typeof final_answer);

  //<-- 回答速度 -->
  const answerTime_sec = calAnserTime(inputTime, quiz_meta);
  console.log("回答速度（sec）", answerTime_sec, typeof answerTime_sec);

  //<-- 回答の答え合わせ -->
  const checkAnswerResult = checkAnswer(quiz_meta, final_answer);
  console.log("回答のチェック:", checkAnswerResult, typeof checkAnswerResult);

  //<-- 埋め込み用の配列 -->
  const addArray = [
    [final_answer, full_name, answerTime_sec, checkAnswerResult],
  ];
  console.log("addArray:", addArray);

  //<-- 算出データを書き込み -->
  sheet.getRange(lastRow, 8, 1, 4).setValues(addArray);

  //<-- TotalDataシートへの書き込み -->
  if (quizNumber !== "Practice") {
    setTotalDataSheet(
      ss,
      quizNumber,
      full_name,
      answerTime_sec,
      checkAnswerResult
    );
  }
}

function getQuizMeta(ss, quizNumber) {
  // クイズのメタシートを取得
  const quiz_meta_sheet = ss.getSheetByName("quiz_meta");

  // metaデータの取得
  const quiz_meta = quiz_meta_sheet.getRange("A2:D10").getValues();

  // [ [ 'Practice', 1, '', '' ],
  // [ 'Question_1', 2, '', '' ],
  // [ 'Question_2', 3, '', '' ],
  // [ 'Question_3', 4, '', '' ],
  // [ 'Question_4', 1, '', '' ],
  // [ 'Question_5', 2, '', '' ],
  // [ 'Question_6', 3, '', '' ],
  // [ 'Question_7', 4, '', '' ],
  // [ 'Question_8', 1, '', '' ] ]

  // リストを辞書に変更
  const quizNumbers = [
    "Practice",
    "Question_1",
    "Question_2",
    "Question_3",
    "Question_4",
    "Question_5",
    "Question_6",
    "Question_7",
    "Question_8",
  ];

  const quizKeys = ["correctAnswer", "startTime", "answerLimit"];

  _quiz_meta = {};
  for (let row = 0; row < quiz_meta.length; row++) {
    _quiz_meta[quizNumbers[row]] = {};
    _quiz_meta[quizNumbers[row]]["correctAnswer"] =
      quiz_meta[row][1].toString();
    _quiz_meta[quizNumbers[row]]["startTime"] = quiz_meta[row][2];
    _quiz_meta[quizNumbers[row]]["answerLimit"] = quiz_meta[row][3];
  }

  // _quiz_meta:
  // console.log(_quiz_meta);
  //  { Practice : { correctAnswer: '1', startTime: '', answerLimit: '' },
  //   Question_1: { correctAnswer: '2', startTime: '', answerLimit: '' },
  //   Question_2: { correctAnswer: '3', startTime: '', answerLimit: '' },
  //   Question_3: { correctAnswer: '4', startTime: '', answerLimit: '' },
  //   Question_4: { correctAnswer: '1', startTime: '', answerLimit: '' },
  //   Question_5: { correctAnswer: '2', startTime: '', answerLimit: '' },
  //   Question_6: { correctAnswer: '3', startTime: '', answerLimit: '' },
  //   Question_7: { correctAnswer: '4', startTime: '', answerLimit: '' },
  //   Question_8: { correctAnswer: '1', startTime: '', answerLimit: '' } }

  return _quiz_meta[quizNumber];
}

function calAnserTime(inputTime, quiz_meta) {
  // 回答開始時間の取得
  let startTime = quiz_meta["startTime"];
  let answerTime_sec;

  if (startTime === "") {
    // まだスタートする前に入力があった場合は、現在時刻をstartTimeとする
    console.info("開始時間を入力前に提出");
    // startTime = new Date();
    answerTime_sec = 0;
  } else {
    // 回答時間の差分を算出
    answerTime_sec = parseInt((inputTime - startTime) / 1000);
  }
  return answerTime_sec;
}

function checkAnswer(quiz_meta, answer) {
  // 回答終了時間の取得
  let answerLimit = quiz_meta["answerLimit"];
  if (answerLimit === "") {
    console.info("回答時間中に提出");
    answerLimit = 0;
  }

  // 正解の選択肢を取得
  const correctAnswer = quiz_meta["correctAnswer"];

  // 回答時間と締め切りの差分を計算
  console.log("回答終了時間:", answerLimit, typeof answerLimit);
  console.log("問題の正解:", correctAnswer, typeof correctAnswer);
  console.log("ユーザーの回答：", answer, typeof answer);

  let result;

  // 正解かどうかを確認する
  if (answer === correctAnswer) {
    // 正答
    if (answerLimit === 0) {
      // 回答締め切り前
      console.log("回答ステータス：", "回答締め切り前, 正解, True");
      result = "True";
    } else {
      // 回答締め切り後
      console.log("回答ステータス：", "回答締め切り後の回答のため、False");
      result = "False";
    }
  } else {
    // 誤答
    console.log("回答ステータス：", "回答締切に関わらず、不正解, False");
    result = "False";
  }

  return result;
}

function setTotalDataSheet(
  ss,
  quizNumber,
  full_name,
  answerTime_sec,
  checkAnswerResult
) {
  // 練習問題以外の全クイズの結果をTotalDataシートに追記する
  const totalDataSheet = ss.getSheetByName("TotalData");
  const addArray = [quizNumber, full_name, answerTime_sec, checkAnswerResult];

  // 書き込み
  totalDataSheet.appendRow(addArray);
}
