function onHomepageOpen(e) {
  var builder = CardService.newCardBuilder();
  var section = CardService.newCardSection()
    .addWidget(CardService.newTextParagraph().setText("カレンダーの予定をクリックすると、Notion連携メニューが表示されます。"));
  builder.addSection(section);
  return builder.build();
}

function onEventOpen(e) {
  return onDefaultHomePageOpen(e);
}

function onDefaultHomePageOpen(e) {
  //Logger.log("onDefaultHomePageOpen: " + JSON.stringify(e));
  const event = getCalendarEvent(e);
  Logger.log("finished getCalendarEvent")
  const url = queryDataFromNotion(e);
  return buildCard(event, url);
}

function buildCard(event, url) {
    Logger.log("buildCard start")
    let cardSection1DecoratedText1Icon1 = CardService.newIconImage()
        .setIcon(CardService.Icon.INVITE);

    let cardSection1DecoratedText1 = CardService.newDecoratedText()
        .setTopLabel('Event title')
        .setText(event.title)
        .setStartIcon(cardSection1DecoratedText1Icon1)
        .setWrapText(true);

    let cardSection1DecoratedText2Icon1 = CardService.newIconImage()
        .setIcon(CardService.Icon.FLIGHT_DEPARTURE);

    let cardSection1DecoratedText2 = CardService.newDecoratedText()
        .setTopLabel('Start time')
        .setText(formatDate(event.startTime))
        .setStartIcon(cardSection1DecoratedText2Icon1)
        .setWrapText(true);

    let cardSection1DecoratedText3Icon1 = CardService.newIconImage()
        .setIcon(CardService.Icon.FLIGHT_ARRIVAL);

    let cardSection1DecoratedText3 = CardService.newDecoratedText()
        .setTopLabel('End time')
        .setText(formatDate(event.endTime))
        .setStartIcon(cardSection1DecoratedText3Icon1)
        .setWrapText(true);

    let cardSection1Divider1 = CardService.newDivider();

    // イベントオブジェクトを文字列化してボタンアクションに渡す準備
    let eventString = JSON.stringify(event); 
    Logger.log("event get");
    let cardSection1ButtonList1 = CardService.newButtonSet();

    if (url) {
      // NotionのURLがすでにある場合
      let cardSection1ButtonList1Button2OpenLink1 = CardService.newOpenLink()
          .setUrl(url);

      let cardSection1ButtonList1Button2 = CardService.newTextButton()
          .setText('OPEN IN NOTION')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOpenLink(cardSection1ButtonList1Button2OpenLink1);

      cardSection1ButtonList1.addButton(cardSection1ButtonList1Button2);
    } else {
      // Notionにまだない場合
      let cardSection1ButtonList1Button1Action1 = CardService.newAction()
          .setFunctionName('sendToNotion')
          .setParameters({event: eventString}); // パラメータ名を明示

      let cardSection1ButtonList1Button1 = CardService.newTextButton()
          .setText('SEND TO NOTION')
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(cardSection1ButtonList1Button1Action1);

      cardSection1ButtonList1.addButton(cardSection1ButtonList1Button1);
    }

    Logger.log("finished botton set");
    let cardSection1 = CardService.newCardSection()
        .addWidget(cardSection1DecoratedText1)
        .addWidget(cardSection1DecoratedText2)
        .addWidget(cardSection1DecoratedText3)
        .addWidget(cardSection1Divider1)
        .addWidget(cardSection1ButtonList1);

    let card = CardService.newCardBuilder()
        .addSection(cardSection1)
        .build();
    Logger.log("finished card set");
    return card;
}

function formatDate(date) {
    if (!date) return "";
    // dateが文字列で来てもオブジェクトで来ても大丈夫なように変換
    const d = new Date(date);
    
    const year = d.getFullYear();
    const month = ('0' + (d.getMonth() + 1)).slice(-2);
    const day = ('0' + d.getDate()).slice(-2);
    const hours = ('0' + d.getHours()).slice(-2);
    const minutes = ('0' + d.getMinutes()).slice(-2);

    return `${year}/${month}/${day} ${hours}:${minutes}`;
}


function getCalendarEvent(e) {
    Logger.log("getCalendarEvent start");
    // e.calendar が存在しない場合（念の為のガード）
    if (!e || !e.calendar) {
      return { title: "No Event Selected" }; 
    }

    let calendarId = e.calendar.calendarId;
    let eventId = e.calendar.id;
    // CalendarAppを使って詳細情報を取得
    let calendar = CalendarApp.getCalendarById(calendarId);
    let event = calendar.getEventById(eventId);

    let extractedEvent = {};
    if (event) {
      extractedEvent['title'] = event.getTitle();
      extractedEvent['startTime'] = event.getStartTime().toISOString(); // JSONで渡すためにISO文字列に変換推奨
      extractedEvent['endTime'] = event.getEndTime().toISOString();
      extractedEvent['creators'] = event.getCreators();
      extractedEvent['location'] = event.getLocation();
      extractedEvent['description'] = event.getDescription();
      extractedEvent['id'] = eventId; // カレンダーのイベントID
    } else {
      Logger.log('Event not found.');
      extractedEvent['title'] = "Event Not Found";
    }
    return extractedEvent;
}

function queryDataFromNotion(e){
  // プロパティが設定されていない場合のエラーハンドリング
  const props = PropertiesService.getScriptProperties();
  const databaseId = props.getProperty('databaseId');
  const notionAPIToken = props.getProperty('notionAPIToken');
  
  if (!databaseId || !notionAPIToken) {
    Logger.log("Notion credentials missing");
    return null;
  }

  const notionAPIEndpoint = 'https://api.notion.com/v1/databases/' + databaseId + '/query';

  let eventId = e.calendar.id;
  // Notionデータベースのプロパティ名が "ID" でテキスト形式であることを前提としています
  const payload = {
    'filter': {
      'property': 'ID',
      'rich_text': {
        'equals': eventId
      }
    }
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' +  notionAPIToken,
      'Content-Type': 'application/json',
      'Notion-Version' : '2022-06-28',
    },
    muteHttpExceptions : true,
    payload: JSON.stringify(payload)
  };

  let result = '';
  try {
      result = UrlFetchApp.fetch(notionAPIEndpoint, options);
      if (result.getResponseCode() !== 200) {
          Logger.log("Notion Query Error: " + result.getContentText());
          return '';
      }
      let parsedResult = JSON.parse(result.getContentText());
      if (!parsedResult.results || parsedResult.results.length === 0) {
          result = '';
      } else if (parsedResult.results[0].url) {
          result = parsedResult.results[0].url;
      }
  } catch (e) {
      Logger.log(e);
  }
  return result;
}

function sendToNotion(e) {
  const props = PropertiesService.getScriptProperties();
  const databaseId = props.getProperty('databaseId');
  const notionAPIToken = props.getProperty('notionAPIToken');
  const notionAPIEndpoint = 'https://api.notion.com/v1/pages/';

  // パラメータからイベント情報を復元
  let event = JSON.parse(e.parameters.event);
  
  const payload = {
    parent: { database_id: databaseId },
    properties: {
      "Name": {
        "title": [{
          "text": {
            "content": event.title || "No Title"
          }
        }]
      },
      "Date": {
        "date": {
          start: event.startTime,
          end : event.endTime
        }
      },
      "Organizer": {
        "rich_text": [{
          "text": {
            "content": String(event.creators)
          }
        }]
      },
      "Location": {
        "rich_text": [{
          "text": {
            "content": event.location || ""
          }
        }]
      },
      "ID": {
        "rich_text": [{
          "text": {
            "content": event.id
          }
        }]
      }
    }
  };

  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' +  notionAPIToken,
      'Content-Type': 'application/json',
      'Notion-Version' : '2022-06-28',
    },
    muteHttpExceptions : true,
    payload: JSON.stringify(payload)
  };

  try {
      let result = UrlFetchApp.fetch(notionAPIEndpoint, options);
      
      if (result.getResponseCode() === 200) {
        // 1. Notionからのレスポンスをパースして、作られたページのURLを取得
        let responseJson = JSON.parse(result.getContentText());
        let newPageUrl = responseJson.url;

        // 2. 新しいURLを使って、カードを再構築する
        // (buildCard関数を再利用します)
        let newCard = buildCard(event, newPageUrl);

        // 3. 通知を出しつつ、現在のカードを新しいカードで上書き(updateCard)する
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText("Notionに保存しました！"))
          .setNavigation(CardService.newNavigation().updateCard(newCard)) // ★ここが重要！
          .build();

      } else {
        throw new Error("Notion API Error: " + result.getContentText());
      }

  } catch (error) {
      Logger.log(error);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("エラーが発生しました: " + error.message))
        .build();
  }
}