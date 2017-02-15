"use strict";
var tslib_1 = require("tslib");
var React = require("react");
var Chat_1 = require("./Chat");
var regExpCard = /\^application\/vnd\.microsoft\.card\./i;
var YOUTUBE_DOMAIN = "youtube.com";
var YOUTUBE_WWW_DOMAIN = "www.youtube.com";
var YOUTUBE_SHORT_DOMAIN = "youtu.be";
var YOUTUBE_WWW_SHORT_DOMAIN = "www.youtu.be";
var VIMEO_DOMAIN = "vimeo.com";
var VIMEO_WWW_DOMAIN = "www.vimeo.com";
exports.queryParams = function (src) {
    return src
        .substr(1)
        .split('&')
        .reduce(function (previous, current) {
        var keyValue = current.split('=');
        previous[decodeURIComponent(keyValue[0])] = decodeURIComponent(keyValue[1]);
        return previous;
    }, {});
};
var queryString = function (query) {
    return Object.keys(query)
        .map(function (key) { return encodeURIComponent(key) + '=' + encodeURIComponent(query[key].toString()); })
        .join('&');
};
var Youtube = function (props) {
    return React.createElement("iframe", { type: "text/html", src: "https://" + YOUTUBE_DOMAIN + "/embed/" + props.embedId + "?" + queryString({
            modestbranding: '1',
            loop: props.loop ? '1' : '0',
            autoplay: props.autoPlay ? '1' : '0'
        }) });
};
var Vimeo = function (props) {
    return React.createElement("iframe", { type: "text/html", src: "https://player." + VIMEO_DOMAIN + "/video/" + props.embedId + "?" + queryString({
            title: '0',
            byline: '0',
            portrait: '0',
            badge: '0',
            autoplay: props.autoPlay ? '1' : '0',
            loop: props.loop ? '1' : '0'
        }) });
};
var Video = function (props) {
    var url = document.createElement('a');
    url.href = props.src;
    var urlQueryParams = exports.queryParams(url.search);
    var pathSegments = url.pathname.substr(1).split('/');
    switch (url.hostname) {
        case YOUTUBE_DOMAIN:
        case YOUTUBE_SHORT_DOMAIN:
        case YOUTUBE_WWW_DOMAIN:
        case YOUTUBE_WWW_SHORT_DOMAIN:
            return React.createElement(Youtube, tslib_1.__assign({ embedId: url.hostname === YOUTUBE_DOMAIN || url.hostname === YOUTUBE_WWW_DOMAIN ? urlQueryParams['v'] : pathSegments[pathSegments.length - 1] }, props));
        case VIMEO_WWW_DOMAIN:
        case VIMEO_DOMAIN:
            return React.createElement(Vimeo, tslib_1.__assign({ embedId: pathSegments[pathSegments.length - 1] }, props));
        default:
            return React.createElement("video", tslib_1.__assign({ controls: true }, props));
    }
};
var Media = function (props) {
    switch (props.type) {
        case 'video':
            return React.createElement(Video, tslib_1.__assign({}, props));
        case 'audio':
            return React.createElement("audio", tslib_1.__assign({ controls: true }, props));
        default:
            return React.createElement("img", tslib_1.__assign({}, props));
    }
};
var mediaType = function (url) {
    return url.slice((url.lastIndexOf(".") - 1 >>> 0) + 2).toLowerCase() == 'gif' ? 'image' : 'video';
};
var title = function (title) { return Chat_1.renderIfNonempty(title, function (title) { return React.createElement("h1", null, title); }); };
var subtitle = function (subtitle) { return Chat_1.renderIfNonempty(subtitle, function (subtitle) { return React.createElement("h2", null, subtitle); }); };
var text = function (text) { return Chat_1.renderIfNonempty(text, function (text) { return React.createElement("p", null, text); }); };
exports.AttachmentView = function (props) {
    if (!props.attachment)
        return;
    var attachment = props.attachment;
    var onCardAction = function (cardAction) { return cardAction &&
        (function (e) {
            props.onCardAction(cardAction.type, cardAction.value);
            e.stopPropagation();
        }); };
    var buttons = function (buttons) { return buttons &&
        React.createElement("ul", { className: "wc-card-buttons" }, buttons.map(function (button, index) { return React.createElement("li", { key: index },
            React.createElement("button", { onClick: onCardAction(button) }, button.title)); })); };
    var attachedImage = function (images) { return images && images.length > 0 &&
        React.createElement(Media, { src: images[0].url, onLoad: props.onImageLoad, onClick: onCardAction(images[0].tap) }); };
    switch (attachment.contentType) {
        case "application/vnd.microsoft.card.hero":
            if (!attachment.content)
                return null;
            return (React.createElement("div", { className: 'wc-card hero', onClick: onCardAction(attachment.content.tap) },
                attachedImage(attachment.content.images),
                title(attachment.content.title),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.thumbnail":
            if (!attachment.content)
                return null;
            return (React.createElement("div", { className: 'wc-card thumbnail', onClick: onCardAction(attachment.content.tap) },
                title(attachment.content.title),
                attachedImage(attachment.content.images),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.video":
            if (!attachment.content || !attachment.content.media || attachment.content.media.length === 0)
                return null;
            return (React.createElement("div", { className: 'wc-card video' },
                React.createElement(Media, { type: 'video', src: attachment.content.media[0].url, onLoad: props.onImageLoad, poster: attachment.content.image && attachment.content.image.url, autoPlay: attachment.content.autostart, loop: attachment.content.autoloop }),
                title(attachment.content.title),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.animation":
            if (!attachment.content || !attachment.content.media || attachment.content.media.length === 0)
                return null;
            return (React.createElement("div", { className: 'wc-card animation' },
                React.createElement(Media, { type: mediaType(attachment.content.media[0].url), src: attachment.content.media[0].url, onLoad: props.onImageLoad, poster: attachment.content.image && attachment.content.image.url, autoPlay: attachment.content.autostart, loop: attachment.content.autoloop }),
                title(attachment.content.title),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.audio":
            if (!attachment.content || !attachment.content.media || attachment.content.media.length === 0)
                return null;
            return (React.createElement("div", { className: 'wc-card audio' },
                React.createElement(Media, { type: 'audio', src: attachment.content.media[0].url, autoPlay: attachment.content.autostart, loop: attachment.content.autoloop }),
                title(attachment.content.title),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.signin":
            if (!attachment.content)
                return null;
            return (React.createElement("div", { className: 'wc-card signin' },
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "application/vnd.microsoft.card.receipt":
            if (!attachment.content)
                return null;
            return (React.createElement("div", { className: 'wc-card receipt', onClick: onCardAction(attachment.content.tap) },
                React.createElement("table", null,
                    React.createElement("thead", null,
                        React.createElement("tr", null,
                            React.createElement("th", { colSpan: 2 }, attachment.content.title)),
                        attachment.content.facts && attachment.content.facts.map(function (fact, i) { return React.createElement("tr", { key: 'fact' + i },
                            React.createElement("th", null, fact.key),
                            React.createElement("th", null, fact.value)); })),
                    React.createElement("tbody", null, attachment.content.items && attachment.content.items.map(function (item, i) {
                        return React.createElement("tr", { key: 'item' + i, onClick: onCardAction(item.tap) },
                            React.createElement("td", null,
                                item.image && React.createElement(Media, { src: item.image.url, onLoad: props.onImageLoad }),
                                Chat_1.renderIfNonempty(item.title, function (title) { return React.createElement("div", { className: "title" }, item.title); }),
                                Chat_1.renderIfNonempty(item.subtitle, function (subtitle) { return React.createElement("div", { className: "subtitle" }, item.subtitle); })),
                            React.createElement("td", null, item.price));
                    })),
                    React.createElement("tfoot", null,
                        Chat_1.renderIfNonempty(attachment.content.tax, function (tax) { return React.createElement("tr", null,
                            React.createElement("td", null, props.format.strings.receiptTax),
                            React.createElement("td", null, attachment.content.tax)); }),
                        Chat_1.renderIfNonempty(attachment.content.total, function (total) { return React.createElement("tr", { className: "total" },
                            React.createElement("td", null, props.format.strings.receiptTotal),
                            React.createElement("td", null, attachment.content.total)); }))),
                buttons(attachment.content.buttons)));
        // Deprecated format for Skype channels. For testing legacy bots in Emulator only.
        case "application/vnd.microsoft.card.flex":
            if (!attachment.content)
                return null;
            return (React.createElement("div", { className: 'wc-card flex' },
                attachedImage(attachment.content.images),
                title(attachment.content.title),
                subtitle(attachment.content.subtitle),
                text(attachment.content.text),
                buttons(attachment.content.buttons)));
        case "image/png":
        case "image/jpg":
        case "image/jpeg":
        case "image/gif":
            return React.createElement(Media, { src: attachment.contentUrl, onLoad: props.onImageLoad });
        case "audio/mpeg":
        case "audio/mp4":
            return React.createElement(Media, { type: 'audio', src: attachment.contentUrl });
        case "video/mp4":
            return React.createElement(Media, { type: 'video', src: attachment.contentUrl, onLoad: props.onImageLoad });
        default:
            var contentType = attachment.contentType;
            var unknown = regExpCard.test(contentType) ? props.format.strings.unknownCard : props.format.strings.unknownFile;
            return React.createElement("span", null, unknown.replace('%1', contentType));
    }
};
//# sourceMappingURL=Attachment.js.map