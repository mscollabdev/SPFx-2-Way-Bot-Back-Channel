"use strict";
var App_1 = require("./App");
exports.App = App_1.App;
var Chat_1 = require("./Chat");
exports.Chat = Chat_1.Chat;
var botframework_directlinejs_1 = require("botframework-directlinejs");
exports.DirectLine = botframework_directlinejs_1.DirectLine;
var Attachment_1 = require("./Attachment");
exports.queryParams = Attachment_1.queryParams;
// below are shims for compatibility with old browsers (IE 10 being the main culprit)
require("core-js/modules/es6.string.starts-with");
require("core-js/modules/es6.array.find");
require("core-js/modules/es6.array.find-index");
//# sourceMappingURL=BotChat.js.map