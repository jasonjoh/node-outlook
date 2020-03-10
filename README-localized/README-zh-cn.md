# Office 365 API 客户端库的 Node.js 包装器

此库提供 Outlook [邮件](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations)、[日历](https://msdn.microsoft.com/office/office365/APi/calendar-rest-operations)和[联系人](https://msdn.microsoft.com/office/office365/APi/contacts-rest-operations) API 的轻量型实现。

有关使用此库的示例应用，请参阅 [Outlook 邮件 API 和 Node.js 入门](https://github.com/jasonjoh/node-tutorial)。

> 此库的原始版本是一个简单的包装器，用于在 Node.js 应用中使用[适用于 Cordova 应用程序的 Microsoft Office 365 API 客户端库](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/)。此版本仍包括旧接口，但是没有对此库的该部分进行进一步的开发。建议应用程序继续使用新接口。

## 所需软件

- [适用于 Cordova 应用程序的 Microsoft Office 365 API 客户端库](https://www.nuget.org/packages/Microsoft.Office365.ClientLib.JS/)（已包含）
- [node-XMLHttpRequest](https://github.com/driverdan/node-XMLHttpRequest)

## 参与

请参阅 [CONTRIBUTING.md](CONTRIBUTING.md)

## 安装

应通过 NPM 完成安装：

```Shell
npm install node-outlook
```

## 用法

安装后，将以下内容添加到源文件：

```js
var outlook = require("node-outlook");
```

### 新接口

新接口记录在简单的[参考](reference/node-outlook.md)中（由 [jsdoc-to-markdown](https://github.com/jsdoc2md/jsdoc-to-markdown) 提供）。

#### 配置

通过`基本`命名空间完成库配置：

- `outlook.base.setApiEndpoint` - 使用它来覆盖 API 终结点。默认值使用 Outlook v1.0 终结点：`https://outlook.office.com/api/v1.0`。
- `outlook.base.setAnchorMailbox` - 将它设置为用户的 SMTP 地址，以使 API 终结点能够有效地路由 API 请求。
- `outlook.base.setPreferredTimeZone` - 使用它指定服务器用于在日历 API 中返回日期/时间值的时区。

#### 执行 API 调用

该库为每个 API 提供了一个命名空间。

- `outlook.mail` - 邮件 API
- `outlook.calendar` - 日历 API
- `outlook.contacts` - 联系人 API

每个命名空间都具有最少的函数（未来将提供更多函数）。命名空间之间的用法类似。例如，以下是在 `outlook.mail` 命名空间中调用 `getMessages` 函数的方法：

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

#### 执行原始 API 调用

如果库未实施可执行所需操作的函数，则可以使用 `outlook.base.makeApiCall` 方法来调用在服务器上实施的任何 API。有关如何使用此方法的示例，请参阅 `outlook.mail`、`outlook.calendar` 或 `outlook.contacts` 命名空间中的任何方法实施。

### 旧接口

> 需要提醒的是，将不再开发旧接口。建议使用新接口。

你可以创建如下所示的 `OutlookServices.Client` 对象：

```js
var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office.com/api/v2.0',
  authHelper.getAccessTokenFn('https://outlook.office.com/', session));
```

其中 `authHelper.getAccessTokenFn` 是一种回调方法，用于提供所需的 OAuth2 访问令牌。

## 版权信息

版权所有 (c) Microsoft。保留所有权利。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

----------
在 Twitter 上通过 [@JasonJohMSFT](https://twitter.com/JasonJohMSFT) 与我联系

关注 [Outlook 开发人员博客](https://blogs.msdn.microsoft.com/exchangedev/)