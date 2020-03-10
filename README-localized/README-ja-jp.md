# Office 365 API クライアント ライブラリ用の Node.js ラッパー

このライブラリでは、Outlook の[メール](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations) API、[予定表](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations) API、[連絡先](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) APIの簡易実装が提供されています。

このライブラリを使用するサンプル アプリについては、「[Getting Started with the Outlook Mail API and Node.js (Outlook Mail API と Node.js を使う)](https://github.com/jasonjoh/node-tutorial)」を参照してください。

> このライブラリの元のバージョンは、[Cordova アプリケーション用 Microsoft Office 365 API クライアント ライブラリ](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/)を Node.js アプリから使用できるようにすることを目的とした簡易ラッパーでした。このバージョンには以前のインターフェイスが今も含まれていますが、このライブラリのその部分の開発は現在は行われていません。今後は、アプリケーションで新しいインターフェイスを使用することをお勧めします。

## 必要なソフトウェア

- [Cordova アプリケーション用 Microsoft Office 365 API クライアント ライブラリ](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/) (付属しています)
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## 投稿

[CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## インストール

インストールは NPM 経由で行う必要があります。

```Shell
npm install node-outlook
```

## 使用方法

インストールしたら、ソース ファイルに次の内容を追加します。

```js
var outlook = require("node-outlook");
```

### 新しいインターフェイス

新しいインターフェイスは、簡単な[リファレンス](reference/node-outlook.md) ([jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown) を使用) で説明されています。

#### 構成

ライブラリの構成は、`base` 名前空間を使用して行います。

- `outlook.base.setApiEndpoint` \- これは、API エンドポイントを上書きするために使用します。既定値では、Outlook v1.0 エンドポイントが使用されます: `https://outlook.office.com/api/v1.0`
- `outlook.base.setAnchorMailbox` \- これはユーザーの SMTP アドレスに設定し、API エンドポイントが API 要求を効率的にルーティングできるようにします。
- `outlook.base.setPreferredTimeZone` \- これは、日付/時刻値を予定表 API で返すためにサーバーが使用するタイム ゾーン指定するために使用します。 

#### API 呼び出しを行う

ライブラリでは、各 API に名前空間があります。

- `outlook.mail` \- メール API
- `outlook.calendar` \- 予定表 API
- `outlook.contacts` \- 連絡先 API

各名前空間には、最低限の関数が含まれています (関数は今後追加される予定です)。使用方法は、各名前空間の間で同様です。たとえば、`getMessages` 関数を `outlook.mail` 名前空間で呼び出す方法は次のようになります。

```js
// Specify an OData query parameters to include in the request
var queryParams = {
  '$select': 'Subject,ReceivedDateTime,From',
  '$orderby': 'ReceivedDateTime desc',
  '$top': 10
};

// Set the API endpoint to use the v2.0 endpoint
outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
// Set the anchor mailbox to the user's SMTP address
outlook.base.setAnchorMailbox(email);

outlook.mail.getMessages({token: token, odataParams: queryParams},
  function(error, result){
    if (error) {
      console.log('getMessages returned an error: ' + error);
    }
    else if (result) {
      console.log('getMessages returned ' + result.value.length + ' messages.');
      result.value.forEach(function(message) {
        console.log('  Subject: ' + message.Subject);
        var from = message.From ? message.From.EmailAddress.Name : "NONE";
        console.log('  From: ' + from);
        console.log('  Received: ' + message.ReceivedDateTime.toString());
      });
    }
  });
```

#### API を直接呼び出す

必要な関数がライブラリに実装されていない場合、`outlook.base.makeApiCall` メソッドを使用すると、サーバーで実装されているいずれの API も呼び出すことができます。このメソッドの使用方法の例については、`outlook.mail` 名前空間、`outlook.calendar` 名前空間、または `outlook.contacts` 名前空間でのいずれかのメソッドの実装を確認できます。

### 以前のインターフェイス

> 以前のインターフェイスの開発は終了している点をご留意ください。新しいインターフェイスを使用することをお勧めします。

`OutlookServices.Client` オブジェクトは次のように作成できます。

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

ここでは、`authHelper.getAccessTokenFn` は必要な OAuth2 アクセス トークンを提供するために実装するコールバック メソッドのことです。

## 著作権

Copyright (c) Microsoft.All rights reserved.

このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

----------
Twitter ([@JasonJohMSFT](https://twitter.com/JasonJohMSFT)) でぜひフォローしてください。

「[Outlook 開発者ブログ](https://blogs.msdn.microsoft.com/exchangedev/)」 をフォローしてください。