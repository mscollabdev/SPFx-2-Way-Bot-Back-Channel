"use strict";
var tslib_1 = require("tslib");
var React = require("react");
var ActivityView_1 = require("./ActivityView");
var Chat_1 = require("./Chat");
var react_redux_1 = require("react-redux");
var HistoryContainer = (function (_super) {
    tslib_1.__extends(HistoryContainer, _super);
    function HistoryContainer(props) {
        var _this = _super.call(this, props) || this;
        _this.scrollToBottom = true;
        _this.resizeListener = function () { return _this.autoscroll(); };
        return _this;
    }
    HistoryContainer.prototype.componentDidMount = function () {
        window.addEventListener('resize', this.resizeListener);
    };
    HistoryContainer.prototype.componentWillUnmount = function () {
        window.removeEventListener('resize', this.resizeListener);
    };
    HistoryContainer.prototype.componentWillUpdate = function () {
        this.scrollToBottom = (Math.abs(this.scrollMe.scrollHeight - this.scrollMe.scrollTop - this.scrollMe.offsetHeight) <= 1);
    };
    HistoryContainer.prototype.componentDidUpdate = function () {
        this.autoscroll();
    };
    HistoryContainer.prototype.autoscroll = function () {
        var vAlignBottomPadding = Math.max(0, measureInnerHeight(this.scrollMe) - this.scrollContent.offsetHeight);
        this.scrollContent.style.marginTop = vAlignBottomPadding + 'px';
        if (this.scrollToBottom)
            this.scrollMe.scrollTop = this.scrollMe.scrollHeight - this.scrollMe.offsetHeight;
    };
    HistoryContainer.prototype.onClickRetry = function (activity) {
        this.props.dispatch({ type: 'Send_Message_Retry', clientActivityId: activity.channelData.clientActivityId });
    };
    HistoryContainer.prototype.onCardAction = function (type, value) {
        switch (type) {
            case "imBack":
                Chat_1.sendMessage(this.props.dispatch, value, this.props.user, this.props.format.locale);
                break;
            case "postBack":
                Chat_1.sendPostBack(this.props.botConnection, value, this.props.user, this.props.format.locale);
                break;
            case "call":
            case "openUrl":
            case "playAudio":
            case "playVideo":
            case "showImage":
            case "downloadFile":
            case "signin":
                window.open(value);
                break;
            default:
                Chat_1.konsole.log("unknown button type", type);
        }
    };
    HistoryContainer.prototype.onSelectActivity = function (activity) {
        this.props.selectedActivitySubject.next({ activity: activity });
    };
    HistoryContainer.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: "wc-message-groups", ref: function (div) { return _this.scrollMe = div; } },
            React.createElement("div", { className: "wc-message-group-content", ref: function (div) { return _this.scrollContent = div; } }, this.props.activities.map(function (activity, index) {
                return React.createElement(WrappedActivity, { key: 'message' + index, activity: activity, showTimestamp: index === _this.props.activities.length - 1 || (index + 1 < _this.props.activities.length && suitableInterval(activity, _this.props.activities[index + 1])), selected: activity === _this.props.selectedActivity, fromMe: activity.from.id === _this.props.user.id, format: _this.props.format, onCardAction: function (type, value) { return _this.onCardAction(type, value); }, onClickActivity: _this.props.selectedActivitySubject && (function () { return _this.onSelectActivity(activity); }), onClickRetry: function (e) {
                        // Since this is a click on an anchor, we need to stop it
                        // from trying to actually follow a (nonexistant) link
                        e.preventDefault();
                        e.stopPropagation();
                        _this.onClickRetry(activity);
                    }, onImageLoad: function () { return _this.autoscroll(); } });
            }))));
    };
    return HistoryContainer;
}(React.Component));
exports.History = react_redux_1.connect(function (state) { return ({
    activities: state.history.activities,
    selectedActivity: state.history.selectedActivity,
    format: state.format,
    user: state.connection.user,
    botConnection: state.connection.botConnection,
    selectedActivitySubject: state.connection.selectedActivity
}); })(HistoryContainer);
var getComputedStyleValues = function (el, stylePropertyNames) {
    var s = window.getComputedStyle(el);
    var result = {};
    stylePropertyNames.forEach(function (name) { return result[name] = parseInt(s.getPropertyValue(name)); });
    return result;
};
var measureInnerHeight = function (el) {
    var paddingTop = 'padding-top', paddingBottom = 'padding-bottom';
    var values = getComputedStyleValues(el, [paddingTop, paddingBottom]);
    return el.offsetHeight - values[paddingTop] - values[paddingBottom];
};
var measureOuterWidth = function (el) {
    var marginLeft = 'margin-left', marginRight = 'margin-right';
    var values = getComputedStyleValues(el, [marginLeft, marginRight]);
    return el.offsetWidth + values[marginLeft] + values[marginRight];
};
var suitableInterval = function (current, next) {
    return Date.parse(next.timestamp) - Date.parse(current.timestamp) > 5 * 60 * 1000;
};
var WrappedActivity = (function (_super) {
    tslib_1.__extends(WrappedActivity, _super);
    function WrappedActivity(props) {
        return _super.call(this, props) || this;
    }
    WrappedActivity.prototype.render = function () {
        var _this = this;
        var timeLine;
        switch (this.props.activity.id) {
            case undefined:
                timeLine = React.createElement("span", null, this.props.format.strings.messageSending);
                break;
            case null:
                timeLine = React.createElement("span", null, this.props.format.strings.messageFailed);
                break;
            case "retry":
                timeLine =
                    React.createElement("span", null,
                        this.props.format.strings.messageFailed,
                        ' ',
                        React.createElement("a", { href: ".", onClick: this.props.onClickRetry }, this.props.format.strings.messageRetry));
                break;
            default:
                var sent = void 0;
                if (this.props.showTimestamp)
                    sent = this.props.format.strings.timeSent.replace('%1', (new Date(this.props.activity.timestamp)).toLocaleTimeString());
                timeLine = React.createElement("span", null,
                    this.props.activity.from.name || this.props.activity.from.id,
                    sent);
                break;
        }
        var who = this.props.fromMe ? 'me' : 'bot';
        var wrapperClassName = [
            'wc-message-wrapper',
            this.props.activity.attachmentLayout || 'list',
            this.props.onClickActivity ? 'clickable' : ''
        ].join(' ');
        var contentClassName = [
            'wc-message-content',
            this.props.selected ? 'selected' : ''
        ].join(' ');
        return (React.createElement("div", { "data-activity-id": this.props.activity.id, className: wrapperClassName, onClick: this.props.onClickActivity },
            React.createElement("div", { className: 'wc-message wc-message-from-' + who, ref: function (div) { return _this.messageDiv = div; } },
                React.createElement("div", { className: contentClassName },
                    React.createElement("svg", { className: "wc-message-callout" },
                        React.createElement("path", { className: "point-left", d: "m0,6 l6 6 v-12 z" }),
                        React.createElement("path", { className: "point-right", d: "m6,6 l-6 6 v-12 z" })),
                    React.createElement(ActivityView_1.ActivityView, { activity: this.props.activity, format: this.props.format, onCardAction: this.props.onCardAction, onImageLoad: this.props.onImageLoad, measureParentHorizontalOverflow: function () { return measureOuterWidth(_this.messageDiv) - _this.messageDiv.offsetParent.offsetWidth; } }))),
            React.createElement("div", { className: 'wc-message-from wc-message-from-' + who }, timeLine)));
    };
    return WrappedActivity;
}(React.Component));
exports.WrappedActivity = WrappedActivity;
//# sourceMappingURL=History.js.map