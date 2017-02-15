"use strict";
var tslib_1 = require("tslib");
var botframework_directlinejs_1 = require("botframework-directlinejs");
var Chat_1 = require("./Chat");
var Strings_1 = require("./Strings");
exports.shell = function (state, action) {
    if (state === void 0) { state = {
        input: '',
        sendTyping: false
    }; }
    switch (action.type) {
        case 'Update_Input':
            return tslib_1.__assign({}, state, { input: action.input });
        case 'Send_Message':
            return tslib_1.__assign({}, state, { input: '' });
        case 'Set_Send_Typing':
            return tslib_1.__assign({}, state, { sendTyping: action.sendTyping });
        default:
            return state;
    }
};
exports.format = function (state, action) {
    if (state === void 0) { state = {
        locale: 'en-us',
        options: {
            showHeader: true
        },
        strings: Strings_1.defaultStrings
    }; }
    switch (action.type) {
        case 'Set_Format_Options':
            return tslib_1.__assign({}, state, { options: action.options });
        case 'Set_Locale':
            return tslib_1.__assign({}, state, { locale: action.locale, strings: Strings_1.strings(action.locale) });
        default:
            return state;
    }
};
exports.connection = function (state, action) {
    if (state === void 0) { state = {
        connectionStatus: botframework_directlinejs_1.ConnectionStatus.Uninitialized,
        botConnection: undefined,
        selectedActivity: undefined,
        user: undefined,
        bot: undefined
    }; }
    switch (action.type) {
        case 'Start_Connection':
            return tslib_1.__assign({}, state, { botConnection: action.botConnection, user: action.user, bot: action.bot, selectedActivity: action.selectedActivity });
        case 'Connection_Change':
            return tslib_1.__assign({}, state, { connectionStatus: action.connectionStatus });
        default:
            return state;
    }
};
exports.history = function (state, action) {
    if (state === void 0) { state = {
        activities: [],
        clientActivityBase: Date.now().toString() + Math.random().toString().substr(1) + '.',
        clientActivityCounter: 0,
        selectedActivity: null
    }; }
    Chat_1.konsole.log("history action", action);
    switch (action.type) {
        case 'Receive_Sent_Message': {
            if (!action.activity.channelData || !action.activity.channelData.clientActivityId) {
                // only postBack messages don't have clientActivityId, and these shouldn't be added to the history
                return state;
            }
            var i = state.activities.findIndex(function (activity) {
                return activity.channelData && activity.channelData.clientActivityId === action.activity.channelData.clientActivityId;
            });
            if (i !== -1) {
                var activity = state.activities[i];
                return tslib_1.__assign({}, state, { activities: state.activities.slice(0, i).concat([
                        action.activity
                    ], state.activities.slice(i + 1)), selectedActivity: state.selectedActivity === activity ? action.activity : state.selectedActivity });
            }
        }
        case 'Receive_Message':
            if (state.activities.find(function (a) { return a.id === action.activity.id; }))
                return state; // don't allow duplicate messages
            return tslib_1.__assign({}, state, { activities: state.activities.filter(function (activity) { return activity.type !== "typing"; }).concat([
                    action.activity
                ], state.activities.filter(function (activity) { return activity.from.id !== action.activity.from.id && activity.type === "typing"; })) });
        case 'Send_Message':
            return tslib_1.__assign({}, state, { activities: state.activities.filter(function (activity) { return activity.type !== "typing"; }).concat([
                    tslib_1.__assign({}, action.activity, { timestamp: (new Date()).toISOString(), channelData: { clientActivityId: state.clientActivityBase + state.clientActivityCounter } })
                ], state.activities.filter(function (activity) { return activity.type === "typing"; })), clientActivityCounter: state.clientActivityCounter + 1 });
        case 'Send_Message_Retry': {
            var activity_1 = state.activities.find(function (activity) {
                return activity.channelData && activity.channelData.clientActivityId === action.clientActivityId;
            });
            var newActivity = activity_1.id === undefined ? activity_1 : tslib_1.__assign({}, activity_1, { id: undefined });
            return tslib_1.__assign({}, state, { activities: state.activities.filter(function (activityT) { return activityT.type !== "typing" && activityT !== activity_1; }).concat([
                    newActivity
                ], state.activities.filter(function (activity) { return activity.type === "typing"; })), selectedActivity: state.selectedActivity === activity_1 ? newActivity : state.selectedActivity });
        }
        case 'Send_Message_Succeed':
        case 'Send_Message_Fail': {
            var i = state.activities.findIndex(function (activity) {
                return activity.channelData && activity.channelData.clientActivityId === action.clientActivityId;
            });
            if (i === -1)
                return state;
            var activity = state.activities[i];
            if (activity.id && activity.id != "retry")
                return state;
            var newActivity = tslib_1.__assign({}, activity, { id: action.type === 'Send_Message_Succeed' ? action.id : null });
            return tslib_1.__assign({}, state, { activities: state.activities.slice(0, i).concat([
                    newActivity
                ], state.activities.slice(i + 1)), clientActivityCounter: state.clientActivityCounter + 1, selectedActivity: state.selectedActivity === activity ? newActivity : state.selectedActivity });
        }
        case 'Show_Typing':
            return tslib_1.__assign({}, state, { activities: state.activities.filter(function (activity) { return activity.type !== "typing"; }).concat(state.activities.filter(function (activity) { return activity.from.id !== action.activity.from.id && activity.type === "typing"; }), [
                    action.activity
                ]) });
        case 'Clear_Typing':
            return tslib_1.__assign({}, state, { activities: state.activities.filter(function (activity) { return activity.id !== action.id; }), selectedActivity: state.selectedActivity && state.selectedActivity.id === action.id ? null : state.selectedActivity });
        case 'Select_Activity':
            if (action.selectedActivity === state.selectedActivity)
                return state;
            return tslib_1.__assign({}, state, { selectedActivity: action.selectedActivity });
        default:
            return state;
    }
};
// Epics - chain actions together with async operations
var redux_1 = require("redux");
var Observable_1 = require("rxjs/Observable");
require("rxjs/add/operator/catch");
require("rxjs/add/operator/delay");
require("rxjs/add/operator/do");
require("rxjs/add/operator/filter");
require("rxjs/add/operator/map");
require("rxjs/add/operator/mapTo");
require("rxjs/add/operator/mergeMap");
require("rxjs/add/operator/throttleTime");
require("rxjs/add/observable/empty");
require("rxjs/add/observable/of");
var sendMessage = function (action$, store) {
    return action$.ofType('Send_Message')
        .map(function (action) {
        var state = store.getState();
        var clientActivityId = state.history.clientActivityBase + (state.history.clientActivityCounter - 1);
        return { type: 'Send_Message_Try', clientActivityId: clientActivityId };
    });
};
var trySendMessage = function (action$, store) {
    return action$.ofType('Send_Message_Try')
        .flatMap(function (action) {
        var state = store.getState();
        var clientActivityId = action.clientActivityId;
        var activity = state.history.activities.find(function (activity) { return activity.channelData && activity.channelData.clientActivityId === clientActivityId; });
        if (!activity) {
            Chat_1.konsole.log("trySendMessage: activity not found");
            return Observable_1.Observable.empty();
        }
        return state.connection.botConnection.postActivity(activity)
            .map(function (id) { return ({ type: 'Send_Message_Succeed', clientActivityId: clientActivityId, id: id }); })
            .catch(function (error) { return Observable_1.Observable.of({ type: 'Send_Message_Fail', clientActivityId: clientActivityId }); });
    });
};
var retrySendMessage = function (action$) {
    return action$.ofType('Send_Message_Retry')
        .map(function (action) { return ({ type: 'Send_Message_Try', clientActivityId: action.clientActivityId }); });
};
var updateSelectedActivity = function (action$, store) {
    return action$.ofType('Send_Message_Succeed', 'Send_Message_Fail', 'Show_Typing', 'Clear_Typing')
        .map(function (action) {
        var state = store.getState();
        if (state.connection.selectedActivity)
            state.connection.selectedActivity.next({ activity: state.history.selectedActivity });
        return { type: null };
    });
};
var showTyping = function (action$) {
    return action$.ofType('Show_Typing')
        .delay(3000)
        .map(function (action) { return ({ type: 'Clear_Typing', id: action.activity.id }); });
};
var sendTyping = function (action$, store) {
    return action$.ofType('Update_Input')
        .map(function (_) { return store.getState(); })
        .filter(function (state) { return state.shell.sendTyping; })
        .throttleTime(3000)
        .do(function (_) { return Chat_1.konsole.log("sending typing"); })
        .flatMap(function (state) {
        return state.connection.botConnection.postActivity({
            type: 'typing',
            from: state.connection.user
        })
            .mapTo({ type: null })
            .catch(function (error) { return Observable_1.Observable.of({ type: null }); });
    });
};
// Now we put it all together into a store with middleware
var redux_2 = require("redux");
var redux_observable_1 = require("redux-observable");
exports.createStore = function () {
    return redux_2.createStore(redux_2.combineReducers({
        shell: exports.shell,
        format: exports.format,
        connection: exports.connection,
        history: exports.history
    }), redux_1.applyMiddleware(redux_observable_1.createEpicMiddleware(redux_observable_1.combineEpics(updateSelectedActivity, sendMessage, trySendMessage, retrySendMessage, showTyping, sendTyping))));
};
//# sourceMappingURL=Store.js.map