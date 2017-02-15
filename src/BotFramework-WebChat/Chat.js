"use strict";
var tslib_1 = require("tslib");
var React = require("react");
var botframework_directlinejs_1 = require("botframework-directlinejs");
var History_1 = require("./History");
var Shell_1 = require("./Shell");
var Store_1 = require("./Store");
var react_redux_1 = require("react-redux");
var Chat = (function (_super) {
    tslib_1.__extends(Chat, _super);
    function Chat(props) {
        var _this = _super.call(this, props) || this;
        _this.store = Store_1.createStore();
        exports.konsole.log("BotChat.Chat props", props);
        _this.store.dispatch({
            type: 'Set_Locale',
            locale: props.locale || window.navigator["userLanguage"] || window.navigator.language || 'en'
        });
        if (props.formatOptions)
            _this.store.dispatch({ type: 'Set_Format_Options', options: props.formatOptions });
        if (props.sendTyping)
            _this.store.dispatch({ type: 'Set_Send_Typing', sendTyping: props.sendTyping });
        return _this;
    }
    Chat.prototype.handleIncomingActivity = function (activity) {
        var state = this.store.getState();
        switch (activity.type) {
            case "message":
                this.store.dispatch({ type: activity.from.id === state.connection.user.id ? 'Receive_Sent_Message' : 'Receive_Message', activity: activity });
                break;
            case "typing":
                if (activity.from.id !== state.connection.user.id)
                    this.store.dispatch({ type: 'Show_Typing', activity: activity });
                break;
        }
    };
    Chat.prototype.componentDidMount = function () {
        var _this = this;
        var props = this.props;
        var botConnection = this.props.directLine
            ? (this.botConnection = new botframework_directlinejs_1.DirectLine(this.props.directLine))
            : this.props.botConnection;
        this.store.dispatch({ type: 'Start_Connection', user: props.user, bot: props.bot, botConnection: botConnection, selectedActivity: props.selectedActivity });
        this.connectionStatusSubscription = botConnection.connectionStatus$.subscribe(function (connectionStatus) {
            return _this.store.dispatch({ type: 'Connection_Change', connectionStatus: connectionStatus });
        });
        this.activitySubscription = botConnection.activity$.subscribe(function (activity) { return _this.handleIncomingActivity(activity); }, function (error) { return exports.konsole.log("activity$ error", error); });
        if (props.selectedActivity) {
            this.selectedActivitySubscription = props.selectedActivity.subscribe(function (activityOrID) {
                _this.store.dispatch({
                    type: 'Select_Activity',
                    selectedActivity: activityOrID.activity || _this.store.getState().history.activities.find(function (activity) { return activity.id === activityOrID.id; })
                });
            });
        }
    };
    Chat.prototype.componentWillUnmount = function () {
        this.connectionStatusSubscription.unsubscribe();
        this.activitySubscription.unsubscribe();
        if (this.selectedActivitySubscription)
            this.selectedActivitySubscription.unsubscribe();
        if (this.botConnection)
            this.botConnection.end();
    };
    Chat.prototype.render = function () {
        var state = this.store.getState();
        exports.konsole.log("BotChat.Chat state", state);
        var header;
        if (state.format.options.showHeader)
            header =
                React.createElement("div", { className: "wc-header" },
                    React.createElement("span", null, state.format.strings.title));
        return (React.createElement(react_redux_1.Provider, { store: this.store },
            React.createElement("div", { className: "wc-chatview-panel" },
                header,
                React.createElement(History_1.History, null),
                React.createElement(Shell_1.Shell, null))));
    };
    return Chat;
}(React.Component));
exports.Chat = Chat;
exports.sendMessage = function (dispatch, text, from, locale) {
    if (!text || typeof text !== 'string' || text.trim().length === 0)
        return;
    dispatch({ type: 'Send_Message', activity: {
            type: "message",
            text: text,
            from: from,
            locale: locale,
            timestamp: (new Date()).toISOString()
        } });
};
exports.sendPostBack = function (botConnection, text, from, locale) {
    botConnection.postActivity({
        type: "message",
        text: text,
        from: from,
        locale: locale
    })
        .subscribe(function (id) {
        exports.konsole.log("success sending postBack", id);
    }, function (error) {
        exports.konsole.log("failed to send postBack", error);
    });
};
var attachmentsFromFiles = function (files) {
    var attachments = [];
    for (var i = 0, numFiles = files.length; i < numFiles; i++) {
        var file = files[i];
        attachments.push({
            contentType: file.type,
            contentUrl: window.URL.createObjectURL(file),
            name: file.name
        });
    }
    return attachments;
};
exports.sendFiles = function (dispatch, files, from, locale) {
    dispatch({ type: 'Send_Message', activity: {
            type: "message",
            attachments: attachmentsFromFiles(files),
            from: from,
            locale: locale
        } });
};
exports.renderIfNonempty = function (value, renderer) {
    if (value !== undefined && value !== null && (typeof value !== 'string' || value.length > 0))
        return renderer(value);
};
exports.konsole = {
    log: function (message) {
        var optionalParams = [];
        for (var _i = 1; _i < arguments.length; _i++) {
            optionalParams[_i - 1] = arguments[_i];
        }
        if (typeof (window) !== 'undefined' && window["botchatDebug"] && message)
            console.log.apply(console, [message].concat(optionalParams));
    }
};
//# sourceMappingURL=Chat.js.map