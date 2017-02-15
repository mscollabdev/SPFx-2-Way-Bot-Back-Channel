"use strict";
var tslib_1 = require("tslib");
var React = require("react");
var Attachment_1 = require("./Attachment");
var Carousel_1 = require("./Carousel");
var FormattedText_1 = require("./FormattedText");
var Attachments = function (props) {
    var attachments = props.attachments, attachmentLayout = props.attachmentLayout, otherProps = tslib_1.__rest(props, ["attachments", "attachmentLayout"]);
    if (!attachments || attachments.length == 0)
        return null;
    return attachmentLayout === 'carousel' ?
        React.createElement(Carousel_1.Carousel, tslib_1.__assign({ attachments: attachments }, otherProps))
        :
            React.createElement("div", { className: "wc-list" }, attachments.map(function (attachment, index) {
                return React.createElement(Attachment_1.AttachmentView, tslib_1.__assign({ key: index, attachment: attachment }, otherProps));
            }));
};
var ActivityView = (function (_super) {
    tslib_1.__extends(ActivityView, _super);
    function ActivityView(props) {
        return _super.call(this, props) || this;
    }
    ActivityView.prototype.shouldComponentUpdate = function (nextProps) {
        return this.props.activity !== nextProps.activity || this.props.format !== nextProps.format;
    };
    ActivityView.prototype.render = function () {
        var _a = this.props, activity = _a.activity, otherProps = tslib_1.__rest(_a, ["activity"]);
        switch (activity.type) {
            case 'message':
                return (React.createElement("div", null,
                    React.createElement(FormattedText_1.FormattedText, { text: activity.text, format: activity.textFormat, onImageLoad: otherProps.onImageLoad }),
                    React.createElement(Attachments, tslib_1.__assign({ attachments: activity.attachments, attachmentLayout: activity.attachmentLayout }, otherProps))));
            case 'typing':
                return React.createElement("div", { className: "wc-typing" });
        }
    };
    return ActivityView;
}(React.Component));
exports.ActivityView = ActivityView;
//# sourceMappingURL=ActivityView.js.map