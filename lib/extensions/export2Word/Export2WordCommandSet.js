var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
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
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import { BaseListViewCommandSet } from '@microsoft/sp-listview-extensibility';
import { Convert2Doc } from './Convert2Doc';
import * as pnp from 'sp-pnp-js';
var LOG_SOURCE = 'Export2WordCommandSet';
var Export2WordCommandSet = /** @class */ (function (_super) {
    __extends(Export2WordCommandSet, _super);
    function Export2WordCommandSet() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    Export2WordCommandSet.prototype.onInit = function () {
        Log.info(LOG_SOURCE, 'Initialized Export2WordCommandSet');
        return Promise.resolve();
    };
    Export2WordCommandSet.prototype.onListViewUpdated = function (event) {
        var export2WordCommand = this.tryGetCommand('Export2Word');
        var listUrl = this.context.pageContext.list.title;
        if (export2WordCommand) {
            // This command should be hidden if selected any rows.
            // export2WordCommand.visible = !(event.selectedRows.length > 0);
            export2WordCommand.visible = (event.selectedRows.length === 1); // && listUrl== "Denník dispečera");
        }
    };
    Export2WordCommandSet.prototype.onExecute = function (event) {
        switch (event.itemId) {
            case 'Export2Word':
                var cnvrt2docx = new Convert2Doc(this.context.spHttpClient, this.context.pageContext.web.absoluteUrl, LOG_SOURCE, this.context.pageContext.list.title);
                event.selectedRows.length == 0 ? cnvrt2docx.createDocument() : this.createDocumentSelectedItems(event, cnvrt2docx);
                break;
            default:
                throw new Error('Unknown command');
        }
    };
    /**
     * Creates the documents for the selected items only
     * @param event
     * @param cnvrt2docx
     */
    Export2WordCommandSet.prototype.returnID = function () {
        return this.properties.ID.toString();
    };
    Export2WordCommandSet.prototype.getUserProperties = function () {
        return __awaiter(this, void 0, void 0, function () {
            var pageUrl, userManager, managerFullName;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        pageUrl = "https://pozfond.sharepoint.com/sites/poolcars";
                        userManager = "";
                        managerFullName = "";
                        // rest api sharepoint user properties
                        /*$.ajax({
                            
                            url: pageUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                            
                            method: "GET",
                    
                            headers: { "Accept": "application/json; odata=verbose" },
                    
                            success: function (data) {
                    
                                //var userProfilePropertyValue = data.d.UserProfileProperties.results.find(KeyValuePair => KeyValuePair.Key === userProfilePropertyName).Value;
                                userProperties= data.d["ExtendedManagers"].results[0].split("|");
                                console.log(userProperties[2]);
                            
                                
                    },
                    
                            error: function (error) {
                    
                                console.log("Error in retriving the user profile property:");
                    
                                console.log(error);
                    
                            }
                    
                        });*/
                        // user properties by jsom
                        return [4 /*yield*/, pnp.sp.profiles.myProperties.get().then(function (result) {
                                var userProperties = result.UserProfileProperties;
                                var userPropertyValues = {};
                                console.log("Manazer");
                                console.log(userManager);
                                console.log(userProperties[14]["Value"]);
                                console.log(userPropertyValues);
                                userManager += userProperties[14]["Value"];
                                //console.log(userProperties);
                                userProperties.forEach(function (property) {
                                    userPropertyValues[property.Key] = property.Value;
                                });
                            })];
                    case 1:
                        // rest api sharepoint user properties
                        /*$.ajax({
                            
                            url: pageUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                            
                            method: "GET",
                    
                            headers: { "Accept": "application/json; odata=verbose" },
                    
                            success: function (data) {
                    
                                //var userProfilePropertyValue = data.d.UserProfileProperties.results.find(KeyValuePair => KeyValuePair.Key === userProfilePropertyName).Value;
                                userProperties= data.d["ExtendedManagers"].results[0].split("|");
                                console.log(userProperties[2]);
                            
                                
                    },
                    
                            error: function (error) {
                    
                                console.log("Error in retriving the user profile property:");
                    
                                console.log(error);
                    
                            }
                    
                        });*/
                        // user properties by jsom
                        _a.sent();
                        if (!(userManager != "")) return [3 /*break*/, 3];
                        return [4 /*yield*/, pnp.sp.profiles.getPropertiesFor(userManager)
                                .then(function (result) {
                                managerFullName += result.UserProfileProperties[4]["Value"] + " " + result.UserProfileProperties[6]["Value"];
                            })];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3: return [2 /*return*/, managerFullName];
                }
            });
        });
    };
    Export2WordCommandSet.prototype.dateConvert = function (dateString) {
        //convert SK datumu na ENG. Pri svk datume 30.6. to zobralo ako 6.30 - invalid date
        var myArray = dateString.split(". ");
        var dateArray = [myArray[0], myArray[1]];
        var year = myArray[2].split(" ")[0];
        var timeArray = myArray[2].split(" ")[1].split(":");
        var myDate = new Date(Number(year), (Number(dateArray[1]) - 1), Number(dateArray[0]), Number(timeArray[0]), Number(timeArray[1]));
        //console.log(myDate.toLocaleString("en-US"));
        return myDate.toLocaleString();
    };
    Export2WordCommandSet.prototype.createDocumentSelectedItems = function (event, cnvrt2docx) {
        return __awaiter(this, void 0, void 0, function () {
            var html, index, values, posadka, spat, menoVodica, spz, hodiny, dni, valuesByName, fieldValueByName, dict, url, nadriadeny, zvysok, cisloZiadanky, d1, d2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        html = '<table>';
                        index = 0;
                        values = [];
                        posadka = "";
                        spat = "";
                        menoVodica = "";
                        spz = "";
                        hodiny = 0;
                        dni = "";
                        valuesByName = [];
                        dict = {};
                        nadriadeny = "";
                        zvysok = 0;
                        cisloZiadanky = "";
                        /*var selectedStr = selected.map(function(item){ // loop all Objects
                            return item.id; */
                        event.selectedRows.forEach(function (i) {
                            //html += `<tr style="height:30px"></tr>`;
                            var isAlternate = index % 2 == 0;
                            i.fields.forEach(function (k) {
                                var value = '';
                                var fieldValue = i.getValue(k);
                                //values.push(i.getValue(k));
                                dict[k.internalName] = i.getValue(k);
                                console.log(i.getValue(k) + ": " + k.internalName);
                                /*switch (k.fieldType) {
                                    case "User":
                                    case "Person or Group":
                                      value = fieldValue && fieldValue.length > 0 ? fieldValue[0].title : '';
                                      break;
                                    case "Lookup":
                                      value = fieldValue && fieldValue.length > 0 ? fieldValue[0].lookupValue : '';
                                      break;
                                    case "TaxonomyFieldType":
                                      value = i.getValue(k).Label;
                                      break;
                                    case "URL":
                                      value = `<a href="${i.getValue(k)}" style="cursor:pointer;">${i.getValue(k)}</a>`;
                                      break;
                                    case "DateTime":
                                      value = new Date(i.getValue(k)).toLocaleString();
                                      //value = new Date(i.getValue(k)).toLocaleString()=="Invalid Date" ? fieldValue :"";
                                      //value = i.getValue(k);
                                      break;
                                    default:
                                      value = i.getValue(k);
                                  }*/
                            });
                            index += 1;
                        });
                        console.log(dict);
                        if (dict["acColPosadka"].length > 0) {
                            dict["acColPosadka"].forEach(function (k) {
                                posadka += k.title + ", ";
                            });
                        }
                        if (dict["acColSpiatocnaCesta"] == "Áno") {
                            spat = "a späť";
                        }
                        if (dict["acColVodic"].length > 0) {
                            dict["acColVodic"].forEach(function (k) {
                                menoVodica += k.title + " ";
                            });
                        }
                        if (dict["acColLookupVozidlo"].length > 0) {
                            dict["acColLookupVozidlo"].forEach(function (k) {
                                spz = k.lookupValue + " ";
                            });
                        }
                        d1 = new Date(this.dateConvert(dict["acColDatumCasOd"]));
                        d2 = new Date(this.dateConvert(dict["acColDatumCasDo"]));
                        //prepocet dni
                        dni += Math.floor((Number(d2) - Number(d1)) / 86400000);
                        //prepocet hodin, ak je recionalne cislo , zaokruhli ho
                        zvysok += ((((Number(d2) - Number(d1)) / 1000) % 86400) / 3600) % 1;
                        if (zvysok == 0) {
                            hodiny += (((Number(d2) - Number(d1)) / 1000) % 86400) / 3600;
                        }
                        else {
                            hodiny += Number(((((Number(d2) - Number(d1)) / 1000) % 86400) / 3600).toFixed(2));
                        }
                        this.properties.ID = dict["ID"].toString();
                        cisloZiadanky += new Date().getFullYear() + "/" + dict["ID"];
                        if (Number(dni) < 1) {
                            dni = "";
                        }
                        return [4 /*yield*/, this.getUserProperties().then(function (properties) {
                                nadriadeny = properties;
                            })];
                    case 1:
                        _a.sent();
                        console.log("nadriadeny" + nadriadeny);
                        console.log(dni, hodiny);
                        html += "<table style=\"border-collapse:collapse;border:none;\">\n    <tbody>\n        <tr>\n            <td colspan=\"2\" rowspan=\"4\" style=\"width: 145.25pt;border-width: 1.5pt 1.5pt 1pt;border-style: solid;border-color: windowtext;border-image: initial;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:2.0pt;'><span style=\"font-size:11px;color:#C00000;\">Organiz&aacute;cia (pe\u010Diatka)</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:3.0pt;'><span style=\"font-size:11px;color:#C00000;\">\u017Diadate\u013E &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><strong><span style=\"font-size:13px;\">&nbsp;</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style='font-size:15px;font-family:\"Calibri\",sans-serif;color:black;'>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; " + dict["acColZiadatelOJ"] + " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></strong></p>\n            </td>\n            <td colspan=\"2\" rowspan=\"3\" style=\"width: 134.7pt;border-top: 1.5pt solid windowtext;border-right: 1.5pt solid windowtext;border-bottom: 1.5pt solid windowtext;border-image: initial;border-left: none;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;text-align:center;'><strong><span style=\"font-size:19px;color:#C00000;\">\u017DIADANKA</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;text-align:center;'><strong><span style=\"font-size:19px;color:#C00000;\">na prepravu</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"color:#C00000;\">&nbsp;</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">os&ocirc;b*</span></strong><span style=\"font-size:11px;color:#C00000;\">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <s>n&aacute;kladu*</s>)</span></p>\n            </td>\n            <td rowspan=\"2\" style=\"width: 148.8pt;border-top: 1.5pt solid windowtext;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 17.1pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:2.0pt;'><span style=\"font-size:11px;color:#C00000;\">\u010C&iacute;slo objedn&aacute;vky \u017Eiadate\u013Ea</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:2.0pt;'><strong><em><span style=\"color:#C00000;\">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span></em></strong></p>\n            </td>\n            <td style=\"height:17.1pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td style=\"height:14.2pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td rowspan=\"2\" style=\"width: 148.8pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 16.85pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:2.0pt;'><span style=\"font-size:11px;color:#C00000;\">\u010C&iacute;slo objedn&aacute;vky &uacute;tvaru</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">dopravy</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:18px;font-family:\"Times New Roman\",serif;'><strong>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; " + cisloZiadanky + "</strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;\">&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span></p>\n            </td>\n            <td style=\"height:16.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"2\" style=\"width: 134.7pt;border-top: none;border-right: none;border-left: none;border-image: initial;border-bottom: 1pt solid windowtext;padding: 0in 3.5pt;height: 0.2in;vertical-align: bottom;\">\n                <h2 style='margin:0in;margin-bottom:.0001pt;font-size:15px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:13px;color:#C00000;\">&nbsp;</span></h2>\n            </td>\n            <td style=\"height:.2in;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 23.7pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Men&aacute; cestuj&uacute;cich*) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style=\"font-size:12px;\">" + posadka + "<span style=\"color:#C00000;\">&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></span></p>\n            </td>\n            <td style=\"height:23.7pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Druh, hmotnos\u0165 a rozmer n&aacute;kladu*) &nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><strong><em><span style=\"font-size:15px;color:#C00000;\">&nbsp;</span></em></strong></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:135.0pt;'><strong><span style=\"font-size:13px;color:#C00000;\">&nbsp;</span></strong></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 18.65pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">De\u0148, hodina a miesto pristavenia*) &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><span style=\"font-size:19px;color:black;\">" + d1.getDate() + "." + (d1.getMonth() + 1) + ".&nbsp;-&nbsp;" + d2.getDate() + "." + (d2.getMonth() + 1) + "." + d2.getFullYear() + "&nbsp;</span></p>\n            </td>\n            <td style=\"height:18.65pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Odkia\u013E &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><span style=\"font-size:15px;color:black;\">" + dict["acColOdkial"] + "-" + dict["acColKam"] + " &nbsp;" + spat + " &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Vodi\u010D sa hl&aacute;si u&nbsp;</span></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"5\" style=\"width: 428.75pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Vozidlo je po\u017Eadovan&eacute; na &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<span style=\"font-size:13px;color:black;\">" + hodiny + "</span> &nbsp; &nbsp;hod&iacute;n</span>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<strong><span style=\"font-size:11px;\">&nbsp; &nbsp; &nbsp;</span></strong><span style=\"font-size:11px;color:#C00000;\"><span style=\"font-size:13px;color:black;\">" + dni + "</span>&nbsp; &nbsp; dni &nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-left:.25in;'><span style=\"font-size:11px;color:#C00000;\">&nbsp;</span></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"3\" style=\"width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 20.4pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">&Uacute;\u010Del jazdy &nbsp;</span><span style=\"font-size:13px;color:black;\">" + dict["Title"] + ",&nbsp;</span></p>\n            </td>\n            <td colspan=\"2\" rowspan=\"2\" style=\"width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 20.4pt;vertical-align: top;\">\n                <h1 style='margin:0in;margin-bottom:.0001pt;text-align:center;font-size:21px;font-family:\"Times New Roman\",serif;font-weight:normal;'><strong><span style=\"font-size:15px;color:#C00000;border:solid windowtext 1.0pt;padding:0in;background:white;\">PR&Iacute;KAZ NA JAZDU</span></strong><span style=\"font-size:15px;color:#C00000;border:solid windowtext 1.0pt;padding:0in;background:white;\">&nbsp; &nbsp;</span><span style=\"font-size:15px;color:#C00000;background:  white;\">&nbsp;&nbsp;</span></h1>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:4.0pt;'><span style=\"font-size:11px;color:#C00000;\">Meno vodi\u010Da &nbsp; &nbsp;&nbsp;</span><span style=\"color:black;\">" + menoVodica + "</span></p>\n            </td>\n            <td style=\"height:20.4pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"3\" style=\"width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 27.8pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Vy&uacute;\u010Dtuje na vrub &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span><strong><span style=\"font-size:15px;\">OKa&Scaron;\u010C</span></strong></p>\n            </td>\n            <td style=\"height:27.8pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"3\" rowspan=\"2\" style=\"width: 185.35pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;margin-top:2.0pt;'><span style=\"font-size:11px;color:#C00000;\">Pozn&aacute;mka \u017Eiadate\u013Ea :&nbsp;</span><span style=\"font-size:13px;color:black;\">" + dict["acColPoznamka"] + ",&nbsp;</span></p></p>\n            </td>\n            <td colspan=\"2\" style=\"width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">Druh vozidla &nbsp; &nbsp;&nbsp;</span><span style=\"font-size:11px;color:black;\">" + dict["Vozidlo_x003a_Druh_x0020_vozidla"] + "</span></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td colspan=\"2\" style=\"width: 243.4pt;border-top: none;border-left: none;border-bottom: 1pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 11.85pt;vertical-align: bottom;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">&Scaron;PZ &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;</span><strong><span style=\"font-size:13px;\">" + spz + "</span></strong></p>\n            </td>\n            <td style=\"height:11.85pt;border:none;\"><br></td>\n        </tr>\n        <tr>\n            <td style=\"width: 92.45pt;border-top: none;border-left: 1.5pt solid windowtext;border-bottom: 1.5pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">D&aacute;tum a podpis&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">\u017Eiadate\u013Ea &nbsp;&nbsp;</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">" + nadriadeny + "</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">" + d1.getDate() + "." + (d1.getMonth() + 1) + "." + d1.getFullYear() + "</span></strong></p>\n            </td>\n            <td colspan=\"2\" style=\"width: 92.9pt;border-top: none;border-left: none;border-bottom: 1.5pt solid windowtext;border-right: 1pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">D&aacute;tum a&nbsp;podpis</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">schva\u013Euj&uacute;ceho</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">Ing. Puchelov&aacute;</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">" + d1.getDate() + "." + (d1.getMonth() + 1) + "." + d1.getFullYear() + "</span></strong></p>\n            </td>\n            <td colspan=\"2\" style=\"width: 243.4pt;border-top: none;border-left: none;border-bottom: 1.5pt solid windowtext;border-right: 1.5pt solid windowtext;padding: 0in 3.5pt;height: 47.25pt;vertical-align: top;\">\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">D&aacute;tum a podpis osoby zodpovednej</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><span style=\"font-size:11px;color:#C00000;\">za autoprev&aacute;dzku</span></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">Peter &Scaron;tetina</span></strong></p>\n                <p style='margin:0in;margin-bottom:.0001pt;font-size:16px;font-family:\"Times New Roman\",serif;'><strong><span style=\"font-size:11px;\">" + d1.getDate() + "." + (d1.getMonth() + 1) + "." + d1.getFullYear() + "</span></strong></p>\n            </td>\n            <td style=\"height:47.25pt;border:none;\"><br></td>\n        </tr>\n    </tbody>\n</table>";
                        console.log("cisloZiadanky - return");
                        console.log(this.returnID());
                        return [4 /*yield*/, cnvrt2docx.generateDocument(html, cisloZiadanky)];
                    case 2:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    __decorate([
        override
    ], Export2WordCommandSet.prototype, "onInit", null);
    __decorate([
        override
    ], Export2WordCommandSet.prototype, "onListViewUpdated", null);
    __decorate([
        override
    ], Export2WordCommandSet.prototype, "onExecute", null);
    return Export2WordCommandSet;
}(BaseListViewCommandSet));
export default Export2WordCommandSet;
//# sourceMappingURL=Export2WordCommandSet.js.map