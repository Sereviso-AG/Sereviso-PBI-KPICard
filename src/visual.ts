/*
*  Power BI Visual CLI
*
*  Copyright (c) Microsoft Corporation
*  All rights reserved.
*  MIT License
*
*  Permission is hereby granted, free of charge, to any person obtaining a copy
*  of this software and associated documentation files (the ""Software""), to deal
*  in the Software without restriction, including without limitation the rights
*  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
*  copies of the Software, and to permit persons to whom the Software is
*  furnished to do so, subject to the following conditions:
*
*  The above copyright notice and this permission notice shall be included in
*  all copies or substantial portions of the Software.
*
*  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
*  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
*  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
*  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
*  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
*  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
*  THE SOFTWARE.
*/

"use strict";
import "core-js/stable";
import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
//import IVisualHost = powerbi.extensibility.IVisual;
import IVisual = powerbi.extensibility.visual.IVisual;
import DataView = powerbi.DataView;
import EnumerateVisualObjectInstancesOptions = powerbi.EnumerateVisualObjectInstancesOptions;
import VisualObjectInstance = powerbi.VisualObjectInstance;
import VisualObjectInstanceEnumeration = powerbi.VisualObjectInstanceEnumeration;
import * as d3 from "d3";
type Selection<T extends d3.BaseType> = d3.Selection<T, any, any, any>;
import { VisualSettings } from "./settings";
import { valueFormatter } from "powerbi-visuals-utils-formattingutils";
import IVisualEventService = powerbi.extensibility.IVisualEventService;

export class Visual implements IVisual {
    eventService: IVisualEventService;
    private events: IVisualEventService;
    //private host: IVisualHost;
    private targetHTML: HTMLElement;
    private kpiName1: HTMLElement;
    private settings: VisualSettings;
    private dataValues: string[];
    private actual: number;
    private target1: number;
    private target2: number;
    private actualDisplay: string;
    private target1Display: string;
    private target2Display: string;
    private target1Arrow: string;
    private target2Arrow: string;
    private kpiName: string;
    private kpiColor: string;
    private kpiTextSize: string;
    private kpiFont: string;
    private actualdisplayUnits: string;
    private actualregionalsetting: string;
    private target1regionalsetting: string;
    private target2regionalsetting: string;
    private actualDecimalPlaces: number;
    private actualColor: string;
    private actualFontSize: string;
    private target1DecimalPlaces: number;
    private target1DisplayUnits: string;
    private target1Color: string;
    private target1FontSize: string;
    private target1Prefix: string;
    private target1Show: string;
    private target1ArrowShow: string;
    private target1PosColor: string;
    private target1NegColor: string;
    private target2DecimalPlaces: number;
    private target2DisplayUnits: string;
    private target2Color: string;
    private target2FontSize: string;
    private target2Prefix: string;
    private target2Show: string;
    private target2ArrowShow: string;
    private target2PosColor: string;
    private target2NegColor: string;
    //private borderLogic: string;
    private goodColor: string;
    private neutralColor: string;
    private badColor: string;
    private neutralRangeStart: number;
    private neutralRangeEnd: number;
    private colorUsing: string;
    private borderSize: number;
    private font: string;
    private unitsAbbr: string;
    private unitsAbbr1: string;
    private unitsAbbr2: string;
    private unitDividerValue: number;
    private unitDividerValue1: number;
    private unitDividerValue2: number;
    private actualFormatValue: number;
    private targetFormatValue1: number;
    private targetFormatValue2: number;
    private locale: string;

    constructor(options: VisualConstructorOptions) {
        this.targetHTML = document.createElement("div");
        options.element.appendChild(this.targetHTML);
        this.kpiName1 = document.createElement("div");
        options.element.appendChild(this.kpiName1);
        this.events = options.host.eventService;
        //Initiating default values
        this.actual = 0;
        this.locale = options.host.locale;
        this.target1 = 0;
        this.target2 = 0;
        this.target1Arrow = "";
        this.target2Arrow = "";
        this.kpiName = "KPI Name";
        this.kpiColor = "#444444";
        this.kpiTextSize = "20";
        this.actualdisplayUnits = "none";
        this.actualregionalsetting = "none";
        this.target1regionalsetting = "none";
        this.target2regionalsetting = "none";
        this.actualDecimalPlaces = 0;
        this.actualColor = "#333333";
        this.actualFontSize = "40";
        this.target1DecimalPlaces = 2;
        this.target1DisplayUnits = "none";
        this.target1Color = "#555";
        this.target1FontSize = "20";
        this.target1Prefix = "";
        this.target1Show = "show";
        this.target1ArrowShow = "show";
        this.target1PosColor = "#5CB85C";
        this.target1NegColor = "#C20000";
        this.target2DecimalPlaces = 2;
        this.target2DisplayUnits = "none";
        this.target2Color = "#555";
        this.target2FontSize = "20";
        this.target2Prefix = "";
        this.target2Show = "show";
        this.target2ArrowShow = "show";
        this.target2PosColor = "#5CB85C";
        this.target2NegColor = "#C20000";
        //this.borderLogic = "negativeisbad";
        this.goodColor = "#5CB85C"; //"#32CD32";
        this.neutralColor = "#F2C80F";
        this.badColor = "#C20000"; //"#FD625E";
        this.neutralRangeStart = -50;
        this.neutralRangeEnd = 50;
        this.colorUsing = "indicator";
        this.borderSize = 15;
        this.font = "Arial";
        this.unitsAbbr = "";
        this.unitsAbbr1 = "";
        this.unitsAbbr2 = "";
        this.unitDividerValue = 1;
        this.unitDividerValue1 = 1;
        this.unitDividerValue2 = 1;
        this.actualFormatValue = 0;
        this.targetFormatValue1 = 0;
        this.targetFormatValue2 = 0;
    }


    public update(options: VisualUpdateOptions) {
        this.events.renderingStarted(options);
        //Get the values from the Format pane.
        this.setObjectProperties(options);
        //Calculate the numbers, number formats and arrows.
        this.calculateKPIValues(options);
        //Generate teh skeleton of the visual and places the values into it
        this.generateVisual();
        //Apply the font, color and border formatting
        this.setVisualstyles();
        //
        this.events.renderingFinished(options);
    }


    public setObjectProperties(options: VisualUpdateOptions) {
        if (options.dataViews[0].metadata.objects) {
            if (options.dataViews[0].metadata.objects["kpiName"] && options.dataViews[0].metadata.objects["kpiName"]["name"]) { this.kpiName = options.dataViews[0].metadata.objects["kpiName"]['name'].toString(); }

            if (options.dataViews[0].metadata.objects["kpiName"] && options.dataViews[0].metadata.objects["kpiName"]["color"]) { this.kpiColor = options.dataViews[0].metadata.objects["kpiName"]["color"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiName"] && options.dataViews[0].metadata.objects["kpiName"]["fontSize"]) { this.kpiTextSize = options.dataViews[0].metadata.objects["kpiName"]["fontSize"].toString(); }

            if (options.dataViews[0].metadata.objects["actual"] && options.dataViews[0].metadata.objects["actual"]["actualdisplayUnits"]) { this.actualdisplayUnits = options.dataViews[0].metadata.objects["actual"]["actualdisplayUnits"].toString(); }
            if (options.dataViews[0].metadata.objects["actual"] && options.dataViews[0].metadata.objects["actual"]["actualDecimalPlaces"]) { this.actualDecimalPlaces = +options.dataViews[0].metadata.objects["actual"]["actualDecimalPlaces"].toString(); }
            if (options.dataViews[0].metadata.objects["actual"] && options.dataViews[0].metadata.objects["actual"]["actualColor"]) { this.actualColor = options.dataViews[0].metadata.objects["actual"]["actualColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["actual"] && options.dataViews[0].metadata.objects["actual"]["fontSize"]) { this.actualFontSize = options.dataViews[0].metadata.objects["actual"]["fontSize"].toString(); }
            if (options.dataViews[0].metadata.objects["actual"] && options.dataViews[0].metadata.objects["actual"]["actualregionalsetting"]) { this.actualregionalsetting = options.dataViews[0].metadata.objects["actual"]["actualregionalsetting"].toString(); }

            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1DisplayUnits"]) { this.target1DisplayUnits = options.dataViews[0].metadata.objects["target1"]["target1DisplayUnits"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1DecimalPlaces"]) { this.target1DecimalPlaces = +options.dataViews[0].metadata.objects["target1"]["target1DecimalPlaces"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1Color"]) { this.target1Color = options.dataViews[0].metadata.objects["target1"]["target1Color"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["fontSize"]) { this.target1FontSize = options.dataViews[0].metadata.objects["target1"]["fontSize"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1Prefix"]) { this.target1Prefix = options.dataViews[0].metadata.objects["target1"]["target1Prefix"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1Arrow"]) { this.target1ArrowShow = options.dataViews[0].metadata.objects["target1"]["target1Arrow"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1Show"]) { this.target1Show = options.dataViews[0].metadata.objects["target1"]["target1Show"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1PosColor"]) { this.target1PosColor = options.dataViews[0].metadata.objects["target1"]["target1PosColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1NegColor"]) { this.target1NegColor = options.dataViews[0].metadata.objects["target1"]["target1NegColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target1"] && options.dataViews[0].metadata.objects["target1"]["target1regionalsetting"]) { this.target1regionalsetting = options.dataViews[0].metadata.objects["target1"]["target1regionalsetting"].toString(); }

            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2DisplayUnits"]) { this.target2DisplayUnits = options.dataViews[0].metadata.objects["target2"]["target2DisplayUnits"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2DecimalPlaces"]) { this.target2DecimalPlaces = +options.dataViews[0].metadata.objects["target2"]["target2DecimalPlaces"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2Color"]) { this.target2Color = options.dataViews[0].metadata.objects["target2"]["target2Color"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["fontSize"]) { this.target2FontSize = options.dataViews[0].metadata.objects["target2"]["fontSize"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2Prefix"]) { this.target2Prefix = options.dataViews[0].metadata.objects["target2"]["target2Prefix"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2Arrow"]) { this.target2ArrowShow = options.dataViews[0].metadata.objects["target2"]["target2Arrow"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2Show"]) { this.target2Show = options.dataViews[0].metadata.objects["target2"]["target2Show"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2PosColor"]) { this.target2PosColor = options.dataViews[0].metadata.objects["target2"]["target2PosColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2NegColor"]) { this.target2NegColor = options.dataViews[0].metadata.objects["target2"]["target2NegColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["target2"] && options.dataViews[0].metadata.objects["target2"]["target2regionalsetting"]) { this.target2regionalsetting = options.dataViews[0].metadata.objects["target2"]["target2regionalsetting"].toString(); }

            //if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["borderLogic"]) { this.borderLogic = options.dataViews[0].metadata.objects["kpiColors"]["borderLogic"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["goodColor"]) { this.goodColor = options.dataViews[0].metadata.objects["kpiColors"]["goodColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["neutralColor"]) { this.neutralColor = options.dataViews[0].metadata.objects["kpiColors"]["neutralColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["badColor"]) { this.badColor = options.dataViews[0].metadata.objects["kpiColors"]["badColor"]["solid"]["color"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["neutralRangeStart"]) { this.neutralRangeStart = +options.dataViews[0].metadata.objects["kpiColors"]["neutralRangeStart"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["neutralRangeEnd"]) { this.neutralRangeEnd = +options.dataViews[0].metadata.objects["kpiColors"]["neutralRangeEnd"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["colorUsing"]) { this.colorUsing = options.dataViews[0].metadata.objects["kpiColors"]["colorUsing"].toString(); }
            if (options.dataViews[0].metadata.objects["kpiColors"] && options.dataViews[0].metadata.objects["kpiColors"]["borderSize"]) { this.borderSize = +options.dataViews[0].metadata.objects["kpiColors"]["borderSize"].toString(); }

            if (options.dataViews[0].metadata.objects["others"] && options.dataViews[0].metadata.objects["others"]["font"]) { this.font = options.dataViews[0].metadata.objects["others"]["font"].toString(); }
        }
    }


    public calculateKPIValues(options: VisualUpdateOptions) {
        //Configure number display format
        if (this.actualdisplayUnits == "none") { this.unitsAbbr = ""; this.unitDividerValue = 1; this.actualFormatValue = 0 }
        else if (this.actualdisplayUnits == "thousands") { this.unitsAbbr = "K"; this.unitDividerValue = 1000; this.actualFormatValue = 1001 }
        else if (this.actualdisplayUnits == "millions") { this.unitsAbbr = "M"; this.unitDividerValue = 1000000; this.actualFormatValue = 1e6 }
        else if (this.actualdisplayUnits == "billions") { this.unitsAbbr = "bn"; this.unitDividerValue = 1000000000; this.actualFormatValue = 1e9 }
        else if (this.actualdisplayUnits == "trillions") { this.unitsAbbr = "T"; this.unitDividerValue = 1000000000000; this.actualFormatValue = 1e12 }
        else { this.unitsAbbr = ""; this.unitDividerValue = 1 }

        if (this.target1DisplayUnits == "none") { this.unitsAbbr1 = ""; this.unitDividerValue1 = 1; this.targetFormatValue1 = 0 }
        else if (this.target1DisplayUnits == "thousands") { this.unitsAbbr1 = "K"; this.unitDividerValue1 = 1000; this.targetFormatValue1 = 1001 }
        else if (this.target1DisplayUnits == "millions") { this.unitsAbbr1 = "M"; this.unitDividerValue1 = 1000000; this.targetFormatValue1 = 1e6 }
        else if (this.target1DisplayUnits == "billions") { this.unitsAbbr1 = "bn"; this.unitDividerValue1 = 1000000000; this.targetFormatValue1 = 1e9 }
        else if (this.target1DisplayUnits == "trillions") { this.unitsAbbr1 = "T"; this.unitDividerValue1 = 1000000000000; this.targetFormatValue1 = 1e12 }
        else { this.unitsAbbr1 = ""; this.unitDividerValue1 = 1 }

        if (this.target2DisplayUnits == "none") { this.unitsAbbr2 = ""; this.unitDividerValue2 = 1; this.targetFormatValue2 = 0 }
        else if (this.target2DisplayUnits == "thousands") { this.unitsAbbr2 = "K"; this.unitDividerValue2 = 1000; this.targetFormatValue2 = 1001 }
        else if (this.target2DisplayUnits == "millions") { this.unitsAbbr2 = "M"; this.unitDividerValue2 = 1000000; this.targetFormatValue2 = 1e6 }
        else if (this.target2DisplayUnits == "billions") { this.unitsAbbr2 = "bn"; this.unitDividerValue2 = 1000000000; this.targetFormatValue2 = 1e9 }
        else if (this.target2DisplayUnits == "trillions") { this.unitsAbbr2 = "T"; this.unitDividerValue2 = 1000000000000; this.targetFormatValue2 = 1e12 }
        else { this.unitsAbbr2 = ""; this.unitDividerValue2 = 1 }

        //Get the numbers from the data
        if (options.dataViews[0].categorical.values[0]) { options.dataViews[0].categorical.values[0].values[0] == null ? this.actual = 0 : this.actual = +options.dataViews[0].categorical.values[0].values[0].toString(); }
        if (options.dataViews[0].categorical.values[1]) { options.dataViews[0].categorical.values[1].values[0] == null ? this.target1 = 0 : this.target1 = +options.dataViews[0].categorical.values[1].values[0].toString(); }
        if (options.dataViews[0].categorical.values[2]) { options.dataViews[0].categorical.values[2].values[0] == null ? this.target2 = 0 : this.target2 = +options.dataViews[0].categorical.values[2].values[0].toString(); }
        
      

        if (this.kpiName) { this.kpiName == 'KPI Name' ? this.kpiName = "KPI Name" : this.kpiName = options.dataViews[0].metadata.objects["kpiName"]['name'].toString(); }
        if (this.target1Prefix) { this.target1Prefix == null ? this.target1Prefix = "" : this.target1Prefix = options.dataViews[0].metadata.objects["target1"]["target1Prefix"].toString(); }
        if (this.target2Prefix) { this.target2Prefix == null ? this.target2Prefix = "" : this.target2Prefix = options.dataViews[0].metadata.objects["target2"]["target2Prefix"].toString(); }
        //Apply the number formatting to the numbers
        
        
        if (options.dataViews[0].categorical
            .values[0] && options.dataViews[0].categorical.values[0].source.format) {
            var actualFormatter = valueFormatter.create({ format: options.dataViews[0].categorical.values[0].source.format, value: this.actualFormatValue, precision: this.actualDecimalPlaces });
            this.actualDisplay = actualFormatter.format(this.actual);
        }
        else {
            this.actualDisplay = (this.actual / this.unitDividerValue).toFixed(this.actualDecimalPlaces).toString() + this.unitsAbbr;
        }

        if (options.dataViews[0].categorical.values[1] && options.dataViews[0].categorical.values[1].source.format) {
            var target1Formatter = valueFormatter.create({ format: options.dataViews[0].categorical.values[1].source.format, value: this.targetFormatValue1, precision: this.target1DecimalPlaces });
            this.target1Display = target1Formatter.format(this.target1);
        }
        else {
            this.target1Display = (this.target1 / this.unitDividerValue1).toFixed(this.target1DecimalPlaces).toString() + this.unitsAbbr1;
        }

        if (options.dataViews[0].categorical.values[2] && options.dataViews[0].categorical.values[2].source.format) {
            var target2Formatter = valueFormatter.create({ format: options.dataViews[0].categorical.values[2].source.format, value: this.targetFormatValue2, precision: this.target2DecimalPlaces });
            this.target2Display = target2Formatter.format(this.target2);
        }
        else {
            this.target2Display = (this.target2 / this.unitDividerValue2).toFixed(this.target2DecimalPlaces).toString() + this.unitsAbbr2;
        }

        //Decide the arrows
        if (this.target1 >= 0) { this.target1Arrow = "up" } else { this.target1Arrow = "down" }
        if (this.target2 >= 0) { this.target2Arrow = "up" } else { this.target2Arrow = "down" }

    }


    public generateVisual() {

        
        function sanitizeString(str) {
            str = str.replace(/[^a-z0-9áéíóúñü \.!@&#^*()+=$%~`,/_-∠°Δπ∞γφ^∨:;?"']/gim, "");
           return str;
        }
        
        this.kpiName = sanitizeString(this.kpiName)
        this.target1Prefix = sanitizeString(this.target1Prefix)
        this.target2Prefix = sanitizeString(this.target2Prefix)  
        
        if (this.target1Show == "show" && this.target2Show == "show") {
            this.targetHTML.insertAdjacentHTML("afterbegin", "<link rel='stylesheet' type='text/css' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>" +
                "<table id=kpitable> <tr id=kpiname> <td colspan=2>" + this.kpiName + "</td> </tr> <tr id=calloutvalue> <td colspan=2>" + this.actualDisplay + "</td> </tr>" +
                "<tr id=variances> </tr> </table>");
        }
        //If either of the targets are hidden
        else {
            this.targetHTML.insertAdjacentHTML("afterbegin", "<link rel='stylesheet' type='text/css' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>" +
                "<table id=kpitable> <tr id=kpiname> <td>" + this.kpiName + "</td> </tr> <tr id=calloutvalue> <td>" + this.actualDisplay + "</td> </tr>" +
                "<tr id=variances> </tr> </table>");
        }

        //Which targets to show
        if (this.target1Show == "show") {
            if (this.target1ArrowShow == "show") {
                document.getElementById('variances').insertAdjacentHTML('beforeend', "<td id=target1>" + this.target1Prefix + " " + this.target1Display + " <i id=arrow1" + this.target1Arrow + " class='fa fa-arrow-" + this.target1Arrow + "'></i> </td>");
            }
            else {
                document.getElementById('variances').insertAdjacentHTML('beforeend', "<td id=target1>" + this.target1Prefix + " " + this.target1Display + " </td>");
            }
        }

        if (this.target2Show == "show") {
            if (this.target2ArrowShow == "show") {
                document.getElementById('variances').insertAdjacentHTML('beforeend', "<td id=target2>" + this.target2Prefix + " " + this.target2Display + " <i id=arrow2" + this.target2Arrow + " class='fa fa-arrow-" + this.target2Arrow + "'></i> </td>");
            }
            else {
                document.getElementById('variances').insertAdjacentHTML('beforeend', "<td id=target2>" + this.target2Prefix + " " + this.target2Display + " </td>");
            }
        }

        if (this.target1Show == "hide" && this.target2Show == "hide") {
            document.getElementById('variances').insertAdjacentHTML('beforeend', "<td> </td>");
        }

    }



    public setVisualstyles() {
        //Get IDs of the HTML components
        var elTable = document.getElementById("kpitable");
        var elKpiName = document.getElementById("kpiname");
        var elCalloutValue = document.getElementById("calloutvalue");
        var elVariances = document.getElementById("variances");
        var elTarget1 = null;
        var elTarget2 = null;
        var borderColor = "#000000";
        var elArrow1Up = null;
        var elArrow1Down = null;
        var elArrow2Up = null;
        var elArrow2Down = null;
        var actualType = "none";
        var target1Type = "none";
        var target2Type = "none";
        if (document.getElementById("arrow1up") != null) { elArrow1Up = document.getElementById("arrow1up"); }
        if (document.getElementById("arrow1down") != null) { elArrow1Down = document.getElementById("arrow1down"); }
        if (document.getElementById("arrow2up") != null) { elArrow2Up = document.getElementById("arrow2up"); }
        if (document.getElementById("arrow2down") != null) { elArrow2Down = document.getElementById("arrow2down"); }
        if (document.getElementById("target1") != null) { elTarget1 = document.getElementById("target1"); }
        if (document.getElementById("target2") != null) { elTarget2 = document.getElementById("target2"); }

        //Get the type of both the targets
        if (this.actual >= this.neutralRangeStart && this.actual <= this.neutralRangeEnd) { actualType = "neutral"; }
        else if (this.actual < this.neutralRangeStart) { actualType = "low"; }
        else if (this.actual > this.neutralRangeEnd) { actualType = "high"; }

        if (this.target1 >= this.neutralRangeStart && this.target1 <= this.neutralRangeEnd) { target1Type = "neutral"; }
        else if (this.target1 < this.neutralRangeStart) { target1Type = "low"; }
        else if (this.target1 > this.neutralRangeEnd) { target1Type = "high"; }

        if (this.target2 >= this.neutralRangeStart && this.target2 <= this.neutralRangeEnd) { target2Type = "neutral"; }
        else if (this.target2 < this.neutralRangeStart) { target2Type = "low"; }
        else if (this.target2 > this.neutralRangeEnd) { target2Type = "high"; }

        //Get border color based on the selected indicator/target type
        if (this.colorUsing == 'indicator') {
            if (actualType == "neutral") { borderColor = this.neutralColor; }
            else if (actualType == "low") { borderColor = this.badColor; }
            else if (actualType == "high") { borderColor = this.goodColor; }
            else { borderColor = "#000"; }
        }
        else if (this.colorUsing == 'target1') {
            if (target1Type == "neutral") { borderColor = this.neutralColor; }
            else if (target1Type == "low") { borderColor = this.badColor; }
            else if (target1Type == "high") { borderColor = this.goodColor; }
            else { borderColor = "#000"; }
        }
        else {
            if (target2Type == "neutral") { borderColor = this.neutralColor; }
            else if (target2Type == "low") { borderColor = this.badColor; }
            else if (target2Type == "high") { borderColor = this.goodColor; }
            else { borderColor == "#000"; }
        }

        //Apply border color, font color, font size and font family.
        elTable.style.border = this.borderSize + "px solid " + borderColor;
        elTable.style.fontFamily = this.font;
        elKpiName.style.color = this.kpiColor;
        elKpiName.style.fontSize = this.kpiTextSize + "px";
        elCalloutValue.style.color = this.actualColor;
        elCalloutValue.style.fontSize = this.actualFontSize + "px";
        if (elTarget1 != null) { elTarget1.style.color = this.target1Color; }
        if (elTarget1 != null) { elTarget1.style.fontSize = this.target1FontSize + "px"; }
        if (elTarget2 != null) { elTarget2.style.color = this.target2Color; }
        if (elTarget2 != null) { elTarget2.style.fontSize = this.target2FontSize + "px"; }

        //Apply color to the target's arrows (Up arrows will always have positive color and down arrows will always have negative color)
        if (elArrow1Up != null) { elArrow1Up.style.color = this.target1PosColor }
        if (elArrow1Down != null) { elArrow1Down.style.color = this.target1NegColor }
        if (elArrow2Up != null) { elArrow2Up.style.color = this.target2PosColor }
        if (elArrow2Down != null) { elArrow2Down.style.color = this.target2NegColor }
    }


    public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
        let objectName = options.objectName;
        let objectEnumeration: VisualObjectInstance[] = [];

        switch (objectName) {
            case 'kpiName':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        name: this.kpiName,
                        color: this.kpiColor,
                        fontSize: this.kpiTextSize,
                    },
                    selector: null
                });
                break;
            case 'actual':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        actualdisplayUnits: this.actualdisplayUnits,
                        actualDecimalPlaces: this.actualDecimalPlaces,
                        actualregionalsetting: this.actualregionalsetting,
                        actualColor: this.actualColor,
                        fontSize: this.actualFontSize,
                    },
                    selector: null
                });
                break;
            case 'target1':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        target1Show: this.target1Show,
                        target1Arrow: this.target1ArrowShow,
                        target1regionalsetting: this.target1regionalsetting,
                        target1Color: this.target1Color,
                        fontSize: this.target1FontSize,
                        target1Prefix: this.target1Prefix,
                        target1DisplayUnits: this.target1DisplayUnits,
                        target1DecimalPlaces: this.target1DecimalPlaces,
                        target1PosColor: this.target1PosColor,
                        target1NegColor: this.target1NegColor,
                    },
                    selector: null
                });
                break;
            case 'target2':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        target2Show: this.target2Show,
                        target2Arrow: this.target2ArrowShow,
                        target2regionalsetting: this.target2regionalsetting,
                        target2Color: this.target2Color,
                        fontSize: this.target2FontSize,
                        target2Prefix: this.target2Prefix,
                        target2DisplayUnits: this.target2DisplayUnits,
                        target2DecimalPlaces: this.target2DecimalPlaces,
                        target2PosColor: this.target2PosColor,
                        target2NegColor: this.target2NegColor,
                    },
                    selector: null
                });
                break;
            case 'kpiColors':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        //borderLogic: this.borderLogic,
                        badColor: this.badColor,
                        neutralColor: this.neutralColor,
                        goodColor: this.goodColor,
                        neutralRangeStart: this.neutralRangeStart,
                        neutralRangeEnd: this.neutralRangeEnd,
                        colorUsing: this.colorUsing,
                        borderSize: this.borderSize,
                    },
                    selector: null
                });
                break;
            case 'others':
                objectEnumeration.push({
                    objectName: objectName,
                    properties: {
                        font: this.font,
                    },
                    selector: null
                });
                break;
        };

        return objectEnumeration;
    }

    public destroy(): void {
        
    }

}