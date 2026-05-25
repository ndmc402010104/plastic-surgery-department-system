/*
========================================
Staff.gs
負責人員名單與簽到單補人
========================================
*/


/*
========================================
取得預設人員名單
========================================
*/

function getDefaultStaffList() {

  return [

    { name: '林煌基', emp: 'M000731', role: 'VS' },
    { name: '鄭舉緒', emp: 'M002022', role: 'VS' },
    { name: '林育賢', emp: 'M006988', role: 'VS' },
    { name: '陳柵君', emp: 'M007505', role: 'VS' },
    { name: '林上熙', emp: 'M009487', role: 'VS' },
    { name: '林培新', emp: 'M011646', role: 'VS' },
    { name: '張耀仁', emp: 'M013078', role: 'VS' },
    { name: '蔡可威', emp: 'M013423', role: 'VS' },

    { name: '李冠臻', emp: 'M013934', role: 'F' },
    { name: '郭曰誠', emp: 'M014534', role: 'F' },

    { name: '馬玉坤', emp: 'M015711', role: 'R6' },
    { name: '石益昇', emp: 'M015081', role: 'R5' },
    { name: '黃正豪', emp: 'M015659', role: 'R4' },
    { name: '林奕凱', emp: 'M016152', role: 'R3' },

    { name: '陳若蘋', emp: 'R014514', role: 'NP' },
    { name: '王姿媖', emp: 'R014055', role: 'NP' },

    { name: '吳明娟', emp: 'A017281', role: '秘書' }

  ];

}


/*
========================================
補入簽到單上固定名單沒有的人
========================================
*/

function addExtraPeopleToMeetingSignSheet() {

  const ss =
    SpreadsheetApp.openById(
      SHEET_ID
    );

  const raw =
    ss.getSheetByName(
      RAW_SHEET_NAME
    );

  const sign =
    ss.getSheetByName(
      SIGN_SHEET_NAME
    );

  const rawValues =
    raw
      .getDataRange()
      .getValues();


  // 清掉動態補人區
  sign
    .getRange(
      'A19:A22'
    )
    .clearContent();

  sign
    .getRange(
      'E16:E22'
    )
    .clearContent();


  if (rawValues.length < 2) {
    return;
  }


  const rawPeople =
    rawValues
      .slice(1)
      .map(row => ({

        name:
          String(
            row[3] || ''
          ).trim(),

        role:
          String(
            row[5] || ''
          ).trim()

      }))
      .filter(p => p.name);


  const fixedText =
    sign
      .getRange(
        'A11:H18'
      )
      .getDisplayValues()
      .flat()
      .join(
        ' '
      );


  const extraPeople =
    rawPeople
      .filter(p =>
        !fixedText.includes(
          p.name
        )
      );


  const rank = {

    'VS': 1,
    'F': 2,
    'Fellow': 2,
    'R6': 3,
    'R5': 4,
    'R4': 5,
    'R3': 6,
    'R2': 7,
    'R1': 8,
    'PGY2': 9,
    'PGY1': 10,
    'PGY' :11,
    'NP': 12,
    'Clerk': 13,
    '秘書': 14

  };


  extraPeople.sort((a, b) =>
    (rank[a.role] || 99)
    -
    (rank[b.role] || 99)
  );


  const extraVS =
    extraPeople
      .filter(p =>
        p.role === 'VS'
      );


  const extraF =
    extraPeople
      .filter(p =>
        p.role === 'F'
        ||
        p.role === 'Fellow'
      );


  const others =
    extraPeople
      .filter(p =>
        p.role !== 'VS'
        &&
        p.role !== 'F'
        &&
        p.role !== 'Fellow'
      );


  const leftList = [

    ...extraVS.map(p =>
      'VS ' + p.name
    ),

    'F 李冠臻',

    'F 郭曰誠',

    ...extraF.map(p =>
      'F ' + p.name
    )

  ];


  const overflow =
    [];


  leftList.forEach((text, index) => {

    if (index < 4) {

      sign
        .getRange(
          19 + index,
          1
        )
        .setValue(
          text
        );

    } else {

      overflow.push(
        text
      );

    }

  });


  const rightList = [

    ...overflow,

    ...others.map(p =>
      p.role
        ? p.role + ' ' + p.name
        : p.name
    )

  ];


  rightList.forEach((text, index) => {

    if (index < 7) {

      sign
        .getRange(
          16 + index,
          5
        )
        .setValue(
          text
        );

    }

  });


  SpreadsheetApp.flush();

}