var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { SPListOperations, SPCore, SPCommonOperations, SPFieldOperations } from 'spfxhelper';
import { Log } from '@microsoft/sp-core-library';
var Convert2Doc = /** @class */ (function () {
    function Convert2Doc(spHttp, webUrl, logSource, listName) {
        this._logSource = undefined;
        this._webURL = undefined;
        this._client = undefined;
        this.listName = undefined;
        this.response = [];
        this.currentViewFields = [];
        this.listFieldDetails = [];
        this.cisloZiadanky = '';
        // Returns the List Operations object
        this._listOperations = undefined;
        // Returns the common operations object
        this._commonOperations = undefined;
        // Returns the field Operations object
        this._fieldOperations = undefined;
        this._logSource = logSource;
        this._webURL = webUrl;
        this._client = spHttp;
        this.listName = listName;
    }
    Object.defineProperty(Convert2Doc.prototype, "listOperations", {
        get: function () {
            if (!this._listOperations) {
                this._listOperations = new SPListOperations(this._client, this._webURL, this._logSource);
            }
            return this._listOperations;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Convert2Doc.prototype, "commonOperations", {
        get: function () {
            if (!this._commonOperations) {
                this._commonOperations = new SPCommonOperations(this._client, this._webURL, this._logSource);
            }
            return this._commonOperations;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Convert2Doc.prototype, "fieldOperations", {
        get: function () {
            if (!this._fieldOperations) {
                this._fieldOperations = new SPFieldOperations(this._client, this._webURL, this._logSource);
            }
            return this._fieldOperations;
        },
        enumerable: true,
        configurable: true
    });
    Object.defineProperty(Convert2Doc.prototype, "createSelectQuery", {
        /**
        * Creates the select query for retreiving the records
        */
        get: function () {
            var select = [];
            var expand = [];
            Log.verbose(this._logSource, "initiating query creation based on the current view fields...");
            //Iterate over each query and create the query
            this.listFieldDetails.forEach(function (i) {
                switch (i.TypeAsString) {
                    case "User":
                    case "Person or Group":
                        select.push(i.InternalName + "/Title");
                        expand.push(i.InternalName);
                        break;
                    case "Lookup":
                        select.push(i.InternalName + "/" + i.LookupField);
                        expand.push(i.InternalName);
                        break;
                    default:
                        select.push(i.InternalName);
                }
            });
            Log.verbose(this._logSource, "Query generated: ?$select=" + select.join(',') + "&$expand=" + expand.join(','));
            return "?$select=" + select.join(',') + "&$expand=" + expand.join(',');
        },
        enumerable: true,
        configurable: true
    });
    /**
     * Creates the document for the export for the current list with current view fields
     */
    Convert2Doc.prototype.createDocument = function () {
        return __awaiter(this, void 0, void 0, function () {
            var items, showQnA;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        Log.verbose(this._logSource, "initiating get of all items in the list...");
                        return [4 /*yield*/, this.getItems()];
                    case 1:
                        items = _a.sent();
                        return [4 /*yield*/, this.validateColumnTypes()];
                    case 2:
                        if (_a.sent()) {
                            showQnA = confirm("QnA format can be printed with the selected view. Do you want to proceed with the QnA format ?\n Press Ok to continue with QnA format, cancel to continue with List Format");
                            // export in the format Q&A
                            showQnA ? this.generateQnAFormat(items) : this.generateTableFormat(items);
                        }
                        else {
                            // SHow in grid format
                            this.generateTableFormat(items);
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns all the items in a list
     * @param listName list name on which query needs to be performed
     */
    Convert2Doc.prototype.getItems = function () {
        return __awaiter(this, void 0, void 0, function () {
            var allItems, e_1;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        allItems = { ok: true, result: [] };
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        // Await to get the response for all the queries
                        return [4 /*yield*/, this.getAllItems()];
                    case 2:
                        // Await to get the response for all the queries
                        _a.sent();
                        Log.verbose(this._logSource, "got all items.");
                        Log.verbose(this._logSource, "collecting all items...");
                        // Iterate over the responses and accumlate all the reveived items
                        this.response.forEach(function (i) {
                            if (i.ok) {
                                allItems.result = allItems.result.concat(i.result);
                            }
                            else {
                                Log.error(_this.listOperations.LogSource, i.error);
                            }
                        });
                        Log.verbose(this._logSource, "items collected with the count " + allItems.result.length);
                        return [3 /*break*/, 5];
                    case 3:
                        e_1 = _a.sent();
                        Log.error(this._logSource, new Error("Error in the method Convert2Doc.getItems()"));
                        Log.error(this._logSource, e_1);
                        return [3 /*break*/, 5];
                    case 4: return [2 /*return*/, Promise.resolve(allItems)];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Recursively gets all the items in the list
     * @param nextLink
     */
    Convert2Doc.prototype.getAllItems = function (nextLink) {
        return __awaiter(this, void 0, void 0, function () {
            var _a, _b, _c, _d, e_2;
            return __generator(this, function (_e) {
                switch (_e.label) {
                    case 0:
                        _e.trys.push([0, 6, , 7]);
                        if (!nextLink) return [3 /*break*/, 2];
                        // Get the next batch of items using the next link
                        Log.verbose(this._logSource, "getting the next set of 5000 records");
                        _b = (_a = this.response).push;
                        return [4 /*yield*/, this.listOperations.getListItemsByNextLink(nextLink)];
                    case 1:
                        _b.apply(_a, [_e.sent()]);
                        Log.verbose(this._logSource, "response received");
                        return [3 /*break*/, 5];
                    case 2:
                        // Get the current view fields with all field details
                        Log.verbose(this._logSource, "retreiving all the fields for the current view...");
                        return [4 /*yield*/, this.getCurrentViewFields()];
                    case 3:
                        _e.sent();
                        Log.verbose(this._logSource, "all fields retreived");
                        // Call the first set of items
                        Log.verbose(this._logSource, "getting the first set of 5000 records");
                        _d = (_c = this.response).push;
                        return [4 /*yield*/, this.listOperations.getListItemsByQuery(this.listName, this.createSelectQuery + "&$top=5000")];
                    case 4:
                        _d.apply(_c, [_e.sent()]);
                        Log.verbose(this._logSource, "response received");
                        _e.label = 5;
                    case 5:
                        // Check if the recent revieved response has the next link
                        if (this.response[this.response.length - 1].nextLink) {
                            this.getAllItems(this.response[this.response.length - 1].nextLink);
                        }
                        return [3 /*break*/, 7];
                    case 6:
                        e_2 = _e.sent();
                        Log.error(this._logSource, new Error("Error in the method Convert2Doc.getAllItems()"));
                        Log.error(this._logSource, e_2);
                        return [3 /*break*/, 7];
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Returns the fields in the current View
     * @param listName
     */
    Convert2Doc.prototype.getCurrentViewFields = function () {
        return __awaiter(this, void 0, void 0, function () {
            var viewId, view, defaultView, viewFields, fieldDetails, orderedFields_1, e_3;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 7, , 8]);
                        viewId = SPCore.getParameterValue(location.href, "viewid");
                        if (!viewId) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.commonOperations.queryGETResquest(this._webURL + "/_api/web/lists/getByTitle('" + this.listName + "')/Views('" + viewId + "')/ViewFields")];
                    case 1:
                        view = _a.sent();
                        this.currentViewFields = view.result["Items"];
                        return [3 /*break*/, 5];
                    case 2: return [4 /*yield*/, this.listOperations.getDefaultView(this.listName)];
                    case 3:
                        defaultView = _a.sent();
                        return [4 /*yield*/, this.fieldOperations.getFieldsByView(this.listName, defaultView.view["Title"])];
                    case 4:
                        viewFields = _a.sent();
                        this.currentViewFields = viewFields.details["Items"];
                        _a.label = 5;
                    case 5: return [4 /*yield*/, this.fieldOperations.getFieldsByList(this.listName)];
                    case 6:
                        fieldDetails = _a.sent();
                        this.listFieldDetails = fieldDetails.details.filter(function (i) { return _this.currentViewFields.indexOf(i.InternalName) > -1; });
                        orderedFields_1 = [];
                        this.currentViewFields.forEach(function (i) {
                            orderedFields_1.push(_this.listFieldDetails.filter(function (j) { return j.InternalName == i; })[0]);
                        });
                        this.listFieldDetails = orderedFields_1;
                        return [3 /*break*/, 8];
                    case 7:
                        e_3 = _a.sent();
                        Log.error(this._logSource, new Error("Error in the method Convert2Doc.getCurrentViewFields()"));
                        Log.error(this._logSource, e_3);
                        return [3 /*break*/, 8];
                    case 8: return [2 /*return*/];
                }
            });
        });
    };
    /**
    * Validates the column types for QnA format
    * @param listName listname
    */
    Convert2Doc.prototype.validateColumnTypes = function () {
        return __awaiter(this, void 0, void 0, function () {
            var isSingleLine, isMultiline, isValidforAnswerMode;
            return __generator(this, function (_a) {
                isSingleLine = false;
                isMultiline = false;
                isValidforAnswerMode = false;
                Log.verbose(this._logSource, "Check if the QnA format cane be created from the current view ?");
                this.listFieldDetails.forEach(function (i) {
                    switch (i.TypeDisplayName.toLowerCase()) {
                        case "single line of text":
                        case "computed":
                            isSingleLine = i.Title === "Title" ? true : false;
                            break;
                        case "multiple lines of text":
                            isMultiline = i.Title === "Answer" ? true : false;
                    }
                });
                if ((this.currentViewFields.length == 2 && isSingleLine && isMultiline)) {
                    isValidforAnswerMode = true;
                }
                Log.verbose(this._logSource, isValidforAnswerMode + ", is the response");
                return [2 /*return*/, Promise.resolve(isValidforAnswerMode)];
            });
        });
    };
    /**
     * Generates the table format for the output
     * @param items
     */
    Convert2Doc.prototype.generateTableFormat = function (items) {
        var _this = this;
        var html = '<table>';
        var index = 0;
        items.result.forEach(function (i) {
            html += "<tr style=\"height:30px\"></tr>";
            var isAlternate = index % 2 == 0;
            _this.listFieldDetails.forEach(function (k) {
                var value = '';
                console.log("vypisujem k:" + k);
                switch (k.TypeAsString) {
                    case "User":
                    case "Person or Group":
                        value = i[k.InternalName]["Title"];
                        break;
                    case "Lookup":
                        value = i[k.InternalName][k.LookupField];
                        break;
                    case "TaxonomyFieldType":
                        value = i[k.InternalName]["Label"];
                        break;
                    case "URL":
                        value = "<a href=\"" + i[k.InternalName]["Url"] + "\" style=\"cursor:pointer;\">" + i[k.InternalName]["Description"] + "</a>";
                        break;
                    case "DateTime":
                        value = new Date(i[k.InternalName]).toLocaleString();
                        break;
                    default:
                        value = i[k.InternalName];
                }
                console.log("Cislo Žiadanky z table format");
                _this.cisloZiadanky = new Date().getFullYear() + "/" + k.InternalName["ID"].toString();
                console.log(_this.cisloZiadanky);
                html += "<tr style=\"background-color:" + (isAlternate ? '#f3f3f3' : '#ffffff') + "\">";
                html += "<td style=\"width:30%; border:" + (isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7') + ";\">" + k.Title + "</td>";
                html += "<td style=\"width:70%;border:" + (isAlternate ? '1px solid #ffffff' : '1px solid #bcb7b7') + ";\">" + value + "</td>";
                html += "</tr>";
            });
            index += 1;
        });
        html = html + "</table>";
        this.generateDocument(html, this.cisloZiadanky);
    };
    /**
     * Generates the QnA format for the output
     * @param items
     */
    Convert2Doc.prototype.generateQnAFormat = function (items) {
        var _this = this;
        var QnA = '';
        items.result.forEach(function (i) {
            _this.currentViewFields.forEach(function (k) {
                if (k.toLowerCase().indexOf('title') > -1) {
                    QnA += "<h3>" + i[k] + "</h3>";
                }
                else {
                    QnA += "<p>" + i[k] + "</p>";
                }
            });
        });
        this.generateDocument(QnA, this.cisloZiadanky);
    };
    /**
    * Generates the document and download it
    * @param sourceHTML
    */
    Convert2Doc.prototype.generateDocument = function (sourceHTML, cisloZiadanky) {
        var headerHTML = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>\n        <head><meta charset='utf-8'><title>" + this.listName + "</title></head><body>";
        var titleHTML = "<h1><center>" + this.listName + "</center></h1><hr></hr>";
        var footerHTML = "</body>My Example</html>";
        var sourceHTML = headerHTML + ("<div id=\"source-html\">" + sourceHTML + "</div>") + footerHTML;
        var source = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(sourceHTML);
        console.log("source");
        console.log(source);
        var fileDownload = document.createElement("a");
        document.body.appendChild(fileDownload);
        fileDownload.href = source;
        fileDownload.download = "Žiadanka" + "/" + cisloZiadanky + ".doc";
        fileDownload.click();
        document.body.removeChild(fileDownload);
    };
    return Convert2Doc;
}());
export { Convert2Doc };
//# sourceMappingURL=Convert2Doc.js.map