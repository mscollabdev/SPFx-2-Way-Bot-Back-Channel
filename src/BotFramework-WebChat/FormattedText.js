"use strict";
var tslib_1 = require("tslib");
var Marked = require("marked");
var React = require("react");
var He = require("he");
exports.FormattedText = function (props) {
    if (!props.text || props.text === '')
        return null;
    switch (props.format) {
        case "plain":
            return renderPlainText(props.text);
        default:
            return renderMarkdown(props.text, props.markdownOptions, props.onImageLoad);
    }
};
var renderPlainText = function (text) {
    var lines = text.replace('\r', '').split('\n');
    var elements = lines.map(function (line, i) { return React.createElement("span", { key: i },
        line,
        React.createElement("br", null)); });
    return React.createElement("span", { className: "format-plain" }, elements);
};
var renderMarkdown = function (text, markdownOptions, onImageLoad) {
    var src = text.replace(/<br\s*\/?>/ig, '\r\n\r\n');
    var options = tslib_1.__assign({ gfm: true, tables: true, breaks: false, pedantic: false, sanitize: false, smartLists: true, silent: false, smartypants: true }, markdownOptions);
    var renderer = options.renderer = new ReactRenderer(options, onImageLoad);
    var elements = renderer.getElements(Marked.parse(src, options));
    /*// debug
    const remaining = renderer.elements.filter(el => !!el);
    if (remaining.length) {
        console.warn(`There were ${remaining.length} unused markdown elements!`);
    }*/
    return React.createElement("span", { className: "format-markdown" }, elements);
};
var ReactRenderer = (function () {
    function ReactRenderer(options, onImageLoad) {
        this.options = options;
        this.onImageLoad = onImageLoad;
        this.key = 0;
        this.elements = [];
    }
    /**
     * We're being sneaky here. Marked is expecting us to render html to text and return that.
     * Instead, we're generating react elements and returning their array indices as strings,
     * which are concatenated by Marked into the final output. We return a stringified index that
     * is {{strongly delimited}}. We must do this because Marked can sometimes leak source text
     * into the stream, interspersed with our ids. This leaked text will be detected later and
     * turned into react elements.
     */
    ReactRenderer.prototype.addElement = function (element) {
        var elementId = this.elements.length;
        this.elements.push(element);
        return "{{" + elementId + "}}";
    };
    /**
     * getElements() reads indices from the input string and populates the return array with
     * corresponding react elements. Marked's lexer/parser/compiler may also leak source text
     * into input string. We detect instances of this and convert them to react elements on-the-fly.
     * Sample input text: "{{87}}{{88}}[{{89}}[{{90}}http://example.com/{{91}}"
     */
    ReactRenderer.prototype.getElements = function (text) {
        var _this = this;
        var elements = new Array();
        var re = /^{{\d+}}/g;
        while (true) {
            var len = text.length;
            // Consume elementIds until string end or a leak sequence is encountered
            text = text.replace(re, function (match) {
                var index = Number(match.match(/\d+/)[0]);
                elements.push(_this.elements[index]);
                _this.elements[index] = null;
                return '';
            });
            if (text.length == 0)
                break;
            // Consume leak sequences until string end or an id sequence is encountered
            var next = text.indexOf('{{');
            while (next > 0) {
                var subst = text.substr(0, next);
                subst = He.unescape(subst);
                elements.push(React.createElement("span", { key: this.key++ }, subst));
                text = text.substr(next);
                next = text.indexOf('{{');
            }
            // Return remainder leak sequence
            if (len == text.length) {
                text = He.unescape(text);
                elements.push(React.createElement("span", { key: this.key++ }, text));
                break;
            }
        }
        return elements.filter(function (el) { return !!el; });
    };
    /// MarkedRenderer overrides
    ReactRenderer.prototype.code = function (code, language) {
        return this.addElement(React.createElement("pre", { key: this.key++ },
            React.createElement("code", null, He.unescape(code))));
    };
    ReactRenderer.prototype.blockquote = function (quote) {
        return this.addElement(React.createElement("blockquote", { key: this.key++ }, this.getElements(quote)));
    };
    ReactRenderer.prototype.html = function (html) {
        return this.addElement(React.createElement("span", { key: this.key++ }, html));
    };
    ReactRenderer.prototype.heading = function (text, level, raw) {
        var HeadingTag = "h" + level;
        return this.addElement(React.createElement(HeadingTag, { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.hr = function () {
        return this.addElement(React.createElement("hr", { key: this.key++ }));
    };
    ReactRenderer.prototype.list = function (body, ordered) {
        var ListTag = ordered ? "ol" : "ul";
        return this.addElement(React.createElement(ListTag, { key: this.key++ }, this.getElements(body)));
    };
    ReactRenderer.prototype.listitem = function (text) {
        return this.addElement(React.createElement("li", { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.paragraph = function (text) {
        return this.addElement(React.createElement("p", { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.table = function (header, body) {
        return this.addElement(React.createElement("table", { key: this.key++ },
            React.createElement("thead", null, this.getElements(header)),
            React.createElement("tbody", null, this.getElements(body))));
    };
    ReactRenderer.prototype.tablerow = function (content) {
        return this.addElement(React.createElement("tr", { key: this.key++ }, this.getElements(content)));
    };
    ReactRenderer.prototype.tablecell = function (content, flags) {
        var CellTag = flags.header ? "th" : "td";
        flags.align = flags.align || "initial";
        var inlineStyle = {
            textAlign: flags.align
        };
        return this.addElement(React.createElement(CellTag, { key: this.key++, style: inlineStyle }, this.getElements(content)));
    };
    ReactRenderer.prototype.strong = function (text) {
        return this.addElement(React.createElement("strong", { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.em = function (text) {
        return this.addElement(React.createElement("em", { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.codespan = function (code) {
        return this.addElement(React.createElement("code", { key: this.key++ }, code));
    };
    ReactRenderer.prototype.br = function () {
        return this.addElement(React.createElement("br", { key: this.key++ }));
    };
    ReactRenderer.prototype.del = function (text) {
        return this.addElement(React.createElement("del", { key: this.key++ }, this.getElements(text)));
    };
    ReactRenderer.prototype.unescapeAndSanitizeLink = function (href) {
        try {
            href = He.unescape(href);
            if (this.options.sanitize) {
                var prot = href.toLowerCase();
                if (!(prot.startsWith('http:') || prot.startsWith('https:'))) {
                    return null;
                }
            }
        }
        catch (e) {
            return null;
        }
        return href;
    };
    ReactRenderer.prototype.link = function (href, title, text) {
        href = this.unescapeAndSanitizeLink(href);
        if (!href)
            return '';
        return this.addElement(React.createElement("a", tslib_1.__assign({ key: this.key++ }, { href: href, title: title, target: '_blank' }), this.getElements(text)));
    };
    ReactRenderer.prototype.image = function (href, title, text) {
        var _this = this;
        href = this.unescapeAndSanitizeLink(href);
        if (!href)
            return '';
        return this.addElement(React.createElement("img", tslib_1.__assign({ key: this.key++, onLoad: function () { return _this.onImageLoad(); } }, { src: href, title: title, alt: text })));
    };
    ReactRenderer.prototype.text = function (text) {
        return this.addElement(React.createElement("span", { key: this.key++ }, He.unescape(text)));
    };
    return ReactRenderer;
}());
//# sourceMappingURL=FormattedText.js.map