import { Version } from '@microsoft/sp-core-library';
import {
  PropertyPaneChoiceGroup,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneLink,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { AppInsights } from "applicationinsights-js";

import { CalloutTriggers } from '@pnp/spfx-property-controls/lib/PropertyFieldHeader';
import { PropertyFieldTextWithCallout } from '@pnp/spfx-property-controls/lib/PropertyFieldTextWithCallout';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';
import { IPickerTerms, PropertyFieldEnterpriseTermPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldEnterpriseTermPicker';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import * as jQuery from 'jquery';
import * as $ from 'jquery';
import { sp, ISearchQuery, SearchResults, SortDirection } from "@pnp/sp/presets/all";
import UIkit from 'uikit';
require("uikit/dist/css/uikit.min.css");
require("uikit/dist/js/uikit.min.js");
import Icons from 'uikit/dist/js/uikit-icons';
import * as moment from "moment";
import { Environment, EnvironmentType, DisplayMode } from '@microsoft/sp-core-library';




export interface ISiliconReefNewsWebpartWebPartProps {
  description: string;
  layout: string;
  list: string;
  type: string;
  poll: string;
  listName: string;
  items: string;
  results: boolean;
  live: boolean;
  sort: string;
  promoted: boolean;
  KQLQuery: string;
  posttype: string;
  uniqueref: string;
  seeall: string;
  numberValue: number;

}

export default class SiliconReefNewsWebpartWebPart extends BaseClientSideWebPart<ISiliconReefNewsWebpartWebPartProps> {

  public render(): void {
    let attach;
    let attach1;
   
   
    $("body").append(`

    <style>
    #workbenchPageContent{max-width:100%}
    .null{visibility:hidden}
    #spPropertyPaneContainer > div > div > div > div > div:nth-child(2) > div > div > div > span,#propertyPaneDescriptionId, #spPropertyPaneContainer > div > div > div > div > div:nth-child(2) button{font-size:20px}
    
    .intro {
      color: #666 !important
  }
  
  
  div[data-sp-feature-tag*="Comments"] {
      display: none
  }
  
  
  .ms-Checkbox {
      padding-top: 20px
  }
  
  .title {
      max-height: 46px;
      overflow: hidden;
  }
  
  .pin {
      display: inline-block;
      background: #FEFEFE;
      border: 2px solid #FAFAFA;
      box-shadow: 0 1px 2px rgba(34, 25, 25, 0.4);
      margin: 0 2px 15px;
      -webkit-column-break-inside: avoid;
      -moz-column-break-inside: avoid;
  
      padding: 15px;
      padding-bottom: 5px;
      background: -webkit-linear-gradient(45deg, #FFF, #F9F9F9);
      opacity: 1;
      -webkit-transition: all .2s ease;
      -moz-transition: all .2s ease;
      -o-transition: all .2s ease;
      transition: all .2s ease;
      width: 100%;
  }
  html{zoom:1}
  @media (min-width: 960px) {
      #columns {
          -webkit-column-count: 2;
          -moz-column-count: 2;
          column-count: 2;
      }
  }
  
  @media (min-width: 1100px) {
      #columns {
          -webkit-column-count: 2;
          -moz-column-count: 2;
          column-count: 2;
      }
  }
  @media (min-width: 2100px) {
     html{zoom:1.2}
  }
  @media (min-width: 2500px) {
      html{zoom:1.3}
   }
  #columns:hover .pin:not(:hover) {
      opacity: 0.7;
  }
  
  
  .icon {
      color: gray !important;
  }
  
  .Emoji {
      height: 32px
  }
  
  #columns {
      -webkit-column-count: 3;
      -webkit-column-gap: 10px;
      -webkit-column-fill: auto;
      -moz-column-count: 3;
      -moz-column-gap: 10px;
      -moz-column-fill: auto;
      column-count: 3;
      column-gap: 15px;
      column-fill: auto;
  }
  
  .edit {
      background-color: transparent;
      border: none;
      box-sizing: border-box;
      display: block;
      margin: 0;
      outline: 0;
      overflow: hidden;
      resize: none;
      white-space: pre;
      width: 100%;
    
      font-size: inherit;
      font-weight: inherit;
      line-height: inherit;
      text-align: inherit;
      color: #333333;
      height: 40px;
  }
  
  .post-module {
      position: relative;
      z-index: 1;
      display: block;
      background: #FFFFFF;
      min-width: 25%;
      height: 340px;
      -webkit-box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      -moz-box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  .post-module:hover,
  .hover,
  .ControlZone-control {
      -webkit-box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
      -moz-box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
      box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
  }
  
  .post-module {
      margin-top: 8px;
      margin-bottom: 10px !important;
  }
  
  .post-module .thumbnail {
      height: 400px;
      overflow: hidden;
  }
  
  .post-module .thumbnail .date {
      position: absolute;
      top: 20px;
      right: 20px;
      z-index: 1;
  
      width: 60px;
      height: 60px;
      padding: 12.5px 0;
      -webkit-border-radius: 100%;
      -moz-border-radius: 100%;
      border-radius: 100%;
  
      font-weight: 700;
      text-align: center;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
  }
  
  .post-module:hover .thumbnail img,
  .hover .thumbnail img {
      -webkit-transform: scale(1.1);
      -moz-transform: scale(1.1);
      transform: scale(1.1);
      opacity: 0.6;
  }
  
  .post-module .thumbnail img {
      display: block;
      width: 120%;
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  .post-module .post-content {
      position: absolute;
      bottom: 0px;
      background: #FFFFFF;
      width: 100%;
      padding: 15px;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
      -webkit-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -moz-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -ms-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -o-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
  }
  
  .post-module .post-content .category {
      position: absolute;
      top: -34px;
      left: 0px;
  
      padding: 10px 15px;
  
      font-size: 14px;
      font-weight: 600;
      text-transform: uppercase;
  }
  
  .post-module .thumbnail .date .day {
      font-size: 18px;
  }
  
  .post-module .thumbnail .date .month {
      font-size: 12px;
      text-transform: uppercase;
  }
  
  .post-module .thumbnail .date {
      background-color: white !important;
      color: #8f92b5 !important;
  }
  
  .post-module .thumbnail .date {
      position: absolute;
      top: 20px;
      right: 20px;
      z-index: 1;
  
      width: 60px;
      height: 60px;
      padding: 12.5px 0;
      -webkit-border-radius: 100%;
      -moz-border-radius: 100%;
      border-radius: 100%;
      color: #ffffff;
      font-weight: 700;
      text-align: center;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
  }
  
  .post-module .thumbnail img {
      display: block;
      width: 120%;
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  h4.title {
    color: rgb(41, 41, 41) !Important;

    line-height: 1.2em !important;
    height: 2.4em !important;
    font-size: 16px !important;
    font-weight: 600 !important;
}
  
  h4,
  h2 {
      font-weight: 600 !Important
  }
  
  .post-module .post-content .category {
      text-transform: none !important;
  }
  
  .card {
      border-radius: 2px
  }
  
  .intro {
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    height: 2.4em;
    padding-top: 0px;
    overflow: hidden;
    font-weight: 400;
    color: #333 !important;
    font-size: 14px;
    font-weight: 400;
    line-height: 1.2em;
}
  
  .post-module .post-content .post-meta {
      margin: 30px 0 0;
      color: #999999;
  }
  
  .post-module .post-content .post-meta {
      margin: 30px 0 0;
      color: #999999;
  }
  
  .postmodulefalse {
      height: 210px !important;
  }
  
  .imagesfalse {
      display: none
  }
  
  .uk-card-title {
      font-size: 14px;
      font-weight: 600 !important;
  
      text-transform: uppercase;
  }
  
  .intro {
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    height: 2.4em;
    padding-top: 0px;
    overflow: hidden;
    font-weight: 400;
    color: #333 !important;
    font-size: 14px;
    font-weight: 400;
    line-height: 1.2em;
}
  
h4.title {
  color: rgb(41, 41, 41) !Important;

  line-height: 1.2em !important;
  height: 2.4em !important;
  font-size: 16px !important;
  font-weight: 600 !important;
}
  
  .uk-label {
      text-align: center;
      font: normal normal normal 13px/15px;
      letter-spacing: 0px;
  
      opacity: 1;
      border-radius: 0px;
      padding: 7px;
  }
  
  .uk-grid>ol {
      list-style: none;
      counter-reset: mycounter;
      padding: 0;
  }
  
  .uk-grid>ol li:before {
      content: counter(mycounter);
      counter-increment: mycounter;
      color: red;
      display: inline-block;
      width: 1em;
      margin-left: -1.5em;
      margin-right: 0.5em;
      font-size: 30px;
      text-align: right;
      direction: rtl
  }
  
  .post-module {
      position: relative;
      z-index: 1;
      display: block;
      background: #FFFFFF;
      min-width: 25%;
      height: 340px;
      -webkit-box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      -moz-box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      box-shadow: 0px 1px 2px 0px rgba(0, 0, 0, 0.15);
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  .post-module:hover,
  .hover,
  .ControlZone-control {
      -webkit-box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
      -moz-box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
      box-shadow: 0px 1px 35px 0px rgba(0, 0, 0, 0.3);
  }
  
  .post-module {
      margin-top: 8px;
      margin-bottom: 10px !important;
  }
  
  .post-module .thumbnail {
      height: 400px;
      overflow: hidden;
  }
  
  .post-module .thumbnail .date {
      position: absolute;
      top: 20px;
      right: 20px;
      z-index: 1;
  
      width: 60px;
      height: 60px;
      padding: 12.5px 0;
      -webkit-border-radius: 100%;
      -moz-border-radius: 100%;
      border-radius: 100%;
      color: #ffffff;
      font-weight: 700;
      text-align: center;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
  }
  
  .post-module:hover .thumbnail img,
  .hover .thumbnail img {
      -webkit-transform: scale(1.1);
      -moz-transform: scale(1.1);
      transform: scale(1.1);
      opacity: 0.6;
  }
  
  .post-module .thumbnail img {
      display: block;
      width: 120%;
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  .post-module .post-content {
      position: absolute;
      bottom: 0px;
      background: #FFFFFF;
      width: 100%;
      padding: 15px;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
      -webkit-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -moz-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -ms-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      -o-transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
      transition: all 0.3s cubic-bezier(0.37, 0.75, 0.61, 1.05) 0s;
  }
  
  .post-module .post-content .category {
      position: absolute;
      top: -34px;
      left: 0px;
  
      padding: 10px 15px;
  
      font-size: 14px;
      font-weight: 600;
      text-transform: uppercase;
  }
  
  .post-module .thumbnail .date .day {
      font-size: 18px;
  }
  
  .post-module .thumbnail .date .month {
      font-size: 12px;
      text-transform: uppercase;
  }
  
  .post-module .thumbnail .date {
      background-color: white !important;
      color: #8f92b5 !important;
  }
  
  .post-module .thumbnail .date {
      position: absolute;
      top: 20px;
      right: 20px;
      z-index: 1;
  
      width: 60px;
      height: 60px;
      padding: 12.5px 0;
      -webkit-border-radius: 100%;
      -moz-border-radius: 100%;
      border-radius: 100%;
      color: #ffffff;
      font-weight: 700;
      text-align: center;
      -webkti-box-sizing: border-box;
      -moz-box-sizing: border-box;
      box-sizing: border-box;
  }
  
  .post-module .thumbnail img {
      display: block;
      width: 120%;
      -webkit-transition: all 0.3s linear 0s;
      -moz-transition: all 0.3s linear 0s;
      -ms-transition: all 0.3s linear 0s;
      -o-transition: all 0.3s linear 0s;
      transition: all 0.3s linear 0s;
  }
  
  h4.title {
    color: rgb(41, 41, 41) !Important;

    line-height: 1.2em !important;
    height: 2.4em !important;
    font-size: 16px !important;
    font-weight: 600 !important;
}
  
  
  
  
  
  
  .component-container {
      background: white
  }
  
  .count {
      float: right;
      padding: 20px;
  }




  .in-slide-container {
    margin-top: 28px;
}

.in-slide-container .in-slideshow::before {
    width: 110px;
    height: 110px;
    margin-bottom: -110px;
    content: "";
    background: url(../img/vilisya-ornament.svg) no-repeat;
    left: 65px;
    top: 18px;
    position: absolute;
}

.in-slide-container .in-slideshow::after {
    width: 110px;
    height: 110px;
    margin-bottom: -110px;
    content: "";
    background: url(../img/vilisya-ornament.svg) no-repeat;
    right: 65px;
    bottom: 51px;
    position: absolute;
}

.in-slide-container .uk-slideshow-items {
    z-index: 1;
}

.in-slide-container .uk-slideshow-items .in-slide-text {
    width: 100%;
    background: #f2f2f2;
    padding: 50px 70px 40px 70px;
    position: relative;
    z-index: 1;
}

.in-slide-container .uk-slideshow-items .in-slide-text h2 {
    color: #fda924;
}

.in-slide-container .uk-slideshow-items .in-slide-image {
    position: absolute;
    top: 0;
    right: 0;
}

.in-slide-container .uk-slideshow-items .in-slide-image img {
    width: 100%;
}

.in-slide-container .uk-slidenav-container {
    position: absolute;
    bottom: 45px;
    left: 70px;
    z-index: 2;
}

  
    </style>`);
    let sfilter;
    const cpn = this.properties.poll;
    const width: number = this.domElement.getBoundingClientRect().width;


    const seeall = this.properties.seeall;
    sp.setup({
      spfxContext: this.context,
    });

    var language;
    if (getQueryStringParameter("Page")) {
      language = getQueryStringParameter("Page").split("/")[4];
    } else { language = document.location.href.split("/")[6]; }

    const nav = sp.web.navigation.topNavigationBar;

    const instanceid = this.context.instanceId;


    let appInsightsKey: String;

    appInsightsKey = "39f70f1c-aeed-4ece-8972-029b37259ace";
    AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });


    const uniqueref = Math.floor(Math.random() * 90000) + 10000;

    sp.setup({
      spfxContext: this.context
    });

    function parseDate(dateStr) {
      var date = dateStr.split('-');
      var day = date[0];
      var month = date[1] - 1; //January = 0
      var year = date[2];
      return new Date(year, month, day);
    }
    var siteurl = this.context.pageContext.web.absoluteUrl;
    var relurl = this.context.pageContext.site.serverRequestPath;

    function getQueryStringParameter(param) {
      if (window.location.href.indexOf("?") > -1) {
        var params = document.URL.split("?")[1].split("&"); //Split Current URL With ? after that &
        var strParams = "";
        for (var i = 0; i < params.length; i = i + 1) { //param,parse with given URL parameter
          var singleParam = params[i].split("=");
          if (singleParam[0] == param) {
            return decodeURIComponent(singleParam[1]); //Decode URL Result
          }
        }
      }
    }
    let desc;
    let desca;
    if (this.properties.description == "") { desca = "null"; desc = this.properties.description; } else {
      desc = this.properties.description;
    }
    this.domElement.innerHTML = '<div class="container"> <div style="padding:15px !important" class="webpart-header ' + desca + '">' + desc + '</div><span id="' + uniqueref + 'seeall" class="right ms-Link ' + desca + '" style="color:#666;float:right;position:relative;bottom:35px;right:20px"></span>'

      + '<div id="' + uniqueref + '" style="" class="uk-grid uk-grid-small"><ol  style="display:none; padding-left:45px" id="' + uniqueref + 'numberedlist"></ol></div></div>';
    const viewpinneda = this.properties.layout;
    var viewpinned;
    if (viewpinneda == undefined) { viewpinned = "uk-width-1-3@m"; } else { viewpinned = viewpinneda; }
    const viewtype = this.properties.type;
    const listName = this.properties.listName;
    const sorta = this.properties.sort;
    var sort;
    sort = "Modified";


    var live = this.properties.live;
    var promoted = this.properties.promoted;
    var nummber = this.properties.numberValue;
    var KQLQuery = this.properties.KQLQuery;
    var campaign;
    if (this.properties.poll == undefined) { campaign = ""; } else { campaign = this.properties.poll; }
    var urlfull = window.location.origin + '/sites/' + window.location.href.split("/")[4];
    var stripparams = urlfull.split("?")[0];

    var livequery;
    var promotedquery;

    var newstypeparam;
    var tagsparam;


    var total;
    if (window.location.href.indexOf("layouts") > -1) {
      total = this.properties.numberValue;
      $(".rippleseeall").hide();
    }
    else {
      total = this.properties.numberValue;
    }
    var campaignfilter;
    const thismonth = new Date(new Date().setDate(new Date().getDate() - 0));
    const thismonthString = thismonth.toISOString();
    let promo;











    var top;
    var topicfilter;
    if (this.properties.poll == undefined || this.properties.poll == "*") { topicfilter = ""; } else { topicfilter = " and OData__TopicHeader eq '" + this.properties.poll + "'"; }
    if (this.properties.numberValue != undefined) { top = this.properties.numberValue; } else { top = 3; }
    var filter = "PromotedState eq '2' " + topicfilter;
    sp.web.lists.getByTitle("Site Pages").items.select("Title", "CanvasContent1", "LayoutWebpartsContent", "FileRef", "BannerImageUrl", "ID", "Description", "OData__TopicHeader", "Modified").filter(filter).orderBy(sort, false).top(top).get().then(results => {




      var content = "";

      var uniqueseeall = "#" + uniqueref + "seeall";
      var seeallfilter = "Title eq '" + cpn + "'";
      sp.web.lists.getByTitle("Site Pages").items.select("Title", "CanvasContent1", "LayoutWebpartsContent", "FileRef", "BannerImageUrl", "ID", "Description", "OData__TopicHeader", "Modified").filter(seeallfilter).get().then(campn => {

        if (campn.length) { sfilter = campn[0].FileRef; } else { sfilter = siteurl + '/_layouts/15/SeeAll.aspx?Page=' + relurl + '%2F&InstanceId=' + instanceid; }

        var seallappend = '<a class="rippleseeall" href="' + sfilter + '" data-interception="propogate" aria-disabled="false">See all</a>';

        jQuery(uniqueseeall).html("");

        jQuery(uniqueseeall).append(seallappend);
      });
      results.forEach(result => {
        var filtertext = "Title eq '" + result.OData__TopicHeader + "'";
        sp.web.lists.getByTitle("Site Pages").items.select("Title", "FileRef", "BannerImageUrl", "ID", "Description", "OData__TopicHeader", "Modified").filter(filtertext).get().then(pages => {

          if (pages.length) { filter = pages[0].FileRef; } else { filter = "/_layouts/15/search.aspx/news?q=" + result.OData__TopicHeader; }

          var e = new Date();





          AppInsights.trackEvent('Post appeared on screen', <any>{
            Site: siteurl,
            Title: result.Title,
            ItemId: result.ID,
            Campaign: result.OData__TopicHeader,





          });
          var words = result.CanvasContent1 + " " + result.LayoutWebpartsContent + " " + result.Description + " " + result.Title + " ";

          var count;
          if (words != null || words != undefined) { count = words.split(/\s+/).length; }
          else { count = 0; }
          var readlength = (Math.round((count - 5) / 265).toString());
          let readtime;
          var numberread = Math.round((count - 5) / 265);
          if (numberread < 1) { readtime = "< 1 minute read"; } else (readtime = readlength + " minute read");

          var d_names = new Array("Sun", "Mon", "Tue",
            "Wed", "Thu", "Fri", "Sat");

          var m_names = new Array("Jan", "Feb", "Mar",
            "Apr", "May", "Jun", "Jul", "Aug", "Sep",
            "Oct", "Nov", "Dec");
          var datetouse;
          if (result.Modified == null) { datetouse = result.Modified; } else { datetouse = result.Modified; }
          var d = new Date(datetouse);
          var curr_day = d.getDay();
          var curr_date = d.getDate();
          var sup = "";
          if (curr_date == 1 || curr_date == 21 || curr_date == 31) {
            sup = "st";
          }
          else if (curr_date == 2 || curr_date == 22) {
            sup = "nd";
          }
          else if (curr_date == 3 || curr_date == 23) {
            sup = "rd";
          }
          else {
            sup = "th";
          }
          var curr_month = d.getMonth();
          var curr_year = d.getFullYear();
          var fulldate = d_names[curr_day] + " " + curr_date + "<SUP>"
            + sup + "</SUP> " + m_names[curr_month];

          var imageurl = result.BannerImageUrl.Url.replace("/_layouts/15/getpreview.ashx?path=%2F", "/");

          var hexcode = "rgb(63, 71, 128) !important";

          if (viewtype == undefined || viewtype == "Tile") {
            content += '<div class="' + viewpinned + '" posttype="' + result.OData__TopicHeader + '" style="margin-bottom:20px">'
              + ' <div class="post-module postmodule uk-card">'
              + '<div class="thumbnail images" style="height:200px"><a data-interception="off" href="' + result.FileRef + '"><img alt="' + result.Title + ' image" style="object-fit: fill;"height="152" src="' + imageurl + '"/></a></div>'
              + '<div class="post-content">'
              + ' <a data-interception="off" class="' + result.OData__TopicHeader + '" style="font-size:12px;font-weight:bold;color:rgba(0,0,0,.8);position:relative;" href="' + filter + '"><span class="' + result.OData__TopicHeader + '">#' + result.OData__TopicHeader + '</span></a></br>'
              + '<a data-interception="off" href="' + result.FileRef + '">'
              + '<h4 class="title" style="font-size:16px;height:42px">' + result.Title + '</h4>'
              + '</a>'
              + '<p class="intro ' + result.Description + '" >' + result.Description + '</p>'
              + '<div class="post-meta" style="font-size:13px; color:rgba(0,0,0,.8)">'

              + '<i class="" aria-hidden="true"></i>' + fulldate + ' </a>'
              + ' <span style="float:right;padding-top:2px"><span> <i class="clock outline icon"></i> ' + readtime + ' </span>'
              + '</span></div></div></div></div></div>';




            attach = "#" + uniqueref;
            jQuery(attach).html("");
            jQuery(attach).append(content);
          } else
            if (viewtype == "Slide") {

              var carouselwrapper = `<div uk-slider><div class=""uk-position-relative uk-visible-toggle uk-light" tabindex="-1" uk-slider>

    <ul id="carouselitems`+ uniqueref + `" class="uk-slider-items uk-child-width-1-2 uk-child-width-1-3@m  uk-grid-small uk-grid-match" style="margin:0px">
    </ul>

    <a style="color:black;    display: inline !important;" class="uk-position-center-left uk-position-small " href="#" uk-slidenav-previous uk-slider-item="previous"></a>
    <a style="color:black;    display: inline !important;" class="uk-position-center-right uk-position-small " href="#" uk-slidenav-next uk-slider-item="next"></a>

  </div>
  <ul class="uk-slider-nav uk-dotnav uk-flex-center uk-margin"></ul></div>`;
              attach = "#" + uniqueref;
              jQuery(attach).html("");
              jQuery(attach).append(carouselwrapper);

              content += '<li posttype="' + result.OData__TopicHeader + '" class="' + viewpinned + '" >'
                + ' <div class="post-module postmodule uk-card" >'
                + '<div class="thumbnail images" style="height:200px"><a data-interception="off" href="' + result.FileRef + '"><img alt="' + result.Title + ' image" style="object-fit: fill;"height="152" src="' + imageurl + '"/></a></div>'
                + '<div class="post-content">'
                + '<a data-interception="off" href="' + result.FileRef + '">'
                + '<h4 class="title" style="font-size:16px;height:42px">' + result.Title + '</h4>'
                + '</a>'
                + '<p class="intro ' + result.Description + '" >' + result.Description + '</p>'
                + '<div class="post-meta" style="font-size:13px; color:rgba(0,0,0,.8)">'
                + ' <a data-interception="off" class="' + result.OData__TopicHeader + '" style="font-size:12px;font-weight:bold;color:rgba(0,0,0,.8);position:relative;bottom:10px" href="' + filter + '"><span class="' + result.OData__TopicHeader + '">#' + result.OData__TopicHeader + '</span></a></br>'
                + '<i class="" aria-hidden="true"></i>' + fulldate + ' </a>'
                + ' <span style="float:right;padding-top:2px"><span> <i class="clock outline icon"></i> ' + readtime + ' </span>'
                + '</span></div></div></div></div></li>';
              attach1 = "#carouselitems" + uniqueref;
              jQuery(attach1).html("");

              jQuery(attach1).append(content);
            }
            else if (viewtype == "ImageSlide") {
              var carouselwrapper1 = `<div style="width:100%;height:400px" class="uk-position-relative uk-visible-toggle uk-light" tabindex="-1" uk-slideshow="min-height: 400; max-height: 400">


    <ul  id="carouselitems`+ uniqueref + `" class="uk-slideshow-items " style="min-height: 400px !Important; height:400px !important">

    </ul>

    <div class="uk-light">
    <a class="uk-position-center-left uk-position-small uk-hidden-hover" href="#" uk-slidenav-previous uk-slideshow-item="previous"></a>
    <a class="uk-position-center-right uk-position-small uk-hidden-hover" href="#" uk-slidenav-next uk-slideshow-item="next"></a>
</div>

</div>
<ul class="uk-slideshow-nav uk-dotnav uk-flex-center uk-margin"></ul>`;
              attach = "#" + uniqueref;
              jQuery(attach).html("");
              jQuery(attach).append(carouselwrapper1);
              content += `<li tabindex="-1" class="uk-transition-active uk-overlay-primary" style="max-height:400px">                    
              <div class="uk-width-1-2 uk-visible@s uk-cover-container uk-height-1-1">
                  <img src="`+imageurl+`" data-src="`+imageurl+`" alt="`+result.Title+` image" width="550" height="400" data-uk-img="`+imageurl+`" data-uk-cover="`+imageurl+`" class="uk-cover" style="width: 550px; height: 420px;">
              </div>
              <div class="uk-position-center-right uk-width-1-1 uk-width-1-2@s">                        
                  <div class="uk-light uk-padding">
                      <div  style="text-transform:none" class="`+result.OData__TopicHeader+` uk-label">`+result.OData__TopicHeader+`</div>
                      <h4 class="">`+result.Title+`</h1>
                      <p class="intro `+ result.Description + `" style="color:white !important;max-width:95%;display: -webkit-box;    -webkit-line-clamp: 2;    -webkit-box-orient: vertical;    line-height: 1.2em;    position:relative;   overflow: hidden;">` + result.Description + `</p>
                      <a href="#" style="text-transform:none" class="uk-button uk-button-text">Read More <span class="uk-margin-small-left uk-icon" data-uk-icon="icon: fa-arrow-right; ratio:0.028"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" width="12.544" height="14.336" data-svg="fa-arrow-right"><path d="M190.5 66.9l22.2-22.2c9.4-9.4 24.6-9.4 33.9 0L441 239c9.4 9.4 9.4 24.6 0 33.9L246.6 467.3c-9.4 9.4-24.6 9.4-33.9 0l-22.2-22.2c-9.5-9.5-9.3-25 .4-34.3L311.4 296H24c-13.3 0-24-10.7-24-24v-32c0-13.3 10.7-24 24-24h287.4L190.9 101.2c-9.8-9.3-10-24.8-.4-34.3z"></path></svg></span></a>
                  </div>
              </div>
          </li>`;
              attach1 = "#carouselitems" + uniqueref;
              jQuery(attach1).html("");

              jQuery(attach1).append(content);
            
            }


            else if (viewtype == "Left") {

              content += `<div posttype="` + result.OData__TopicHeader + `" style="margin-left:15px; margin-bottom:10px" class="uk-width-1-1@m uk-card uk-card-default uk-grid-collapse   uk-grid uk-grid-small"  uk-grid>
      <div style="height:170px" class="post-module uk-card-media-left uk-cover-container uk-width-1-4@m">
      <a data-interception="off" href="`+ result.FileRef + `">
       <img class="thumbnail image" style="max-height:170px" src="`+ imageurl + `" alt="` + result.Title + ` image" uk-cover></a>

      </div>
      <div style="height:185px" class="uk-width-expand@m">
          <div style="padding-top:15px;padding-left:20px" class="uk-width-1-1@m">

          <a data-interception="off" style="font-size:11px;font-weight:bold;color:rgba(0,0,0,.8)" href="`+ filter + `"><span style="" class="` + result.OData__TopicHeader + `">#` + result.OData__TopicHeader + `</span></a>
          <a data-interception="off" href="`+ result.FileRef + `">
          <h4 style="-webkit-box;    -webkit-line-clamp: 2;    -webkit-box-orient: vertical;    line-height: 1.2em;     font-size:16px !Important;max-width:90%; margin-bottom:10px;overflow:hidden" class="uk-card-title title">`+ result.Title + `</h4></a>

              <p class="intro `+ result.Description + `" style="max-width:95%;display: -webkit-box;    -webkit-line-clamp: 2;    -webkit-box-orient: vertical;    line-height: 1.2em;    position:relative;   overflow: hidden;">` + result.Description + `</p>
              <div class="post-meta" style="max-width:80%;font-size:13px;position:relative;top:-2px; color:rgba(0,0,0,.8)">

<span class="uk-position-left"> `+ fulldate + ` &nbsp;&nbsp; <i class="clock outline icon"></i>  ` + readtime + `</span>

   </div>
          </div>
      </div>
  </div>`;


              attach = "#" + uniqueref;
              jQuery(attach).html("");

              jQuery(attach).append(content);


            }
            else if (viewtype == "List") {

              content += `
  <li posttype="`+ result.OData__TopicHeader + `" class="uk-width-1-1@m " style="max-width:100%">
        <span style=" margin-left: -2em; text-indent: 2em; position:relative;bottom:5px;left:35px;padding-bottom:10px;padding-top:5px;line-height:22px;font-size:15px;padding-bottom:15px !important;min-width:100%; ">
        <a  style="color: rgb(41,41,41) !Important;line-height: 25px;    height: 55px !important;    font-size: 15px !important;    font-weight: 600;" data-interception="off" href="`+ result.FileRef + `">` + result.Title + ` </a> </span>
        <hr style="position:relative;right:2em;margin-bottom:5px;margin-top:5px" class="uk-width-1-1@m uk-divider-icon"></li>


`;


              attach = "#" + uniqueref;
              jQuery("#" + uniqueref + "numberedlist").html("");
              jQuery("#" + uniqueref + "numberedlist").append(content);
              jQuery("#" + uniqueref + "numberedlist").show();

            }
        });
      });
    });










  }


  private lists: IPropertyPaneDropdownOption[];

  private items: IPropertyPaneDropdownOption[];
  private thisdropitems: IPropertyPaneDropdownOption[];
  private listsDropdownDisabled: boolean = true;

  private loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    sp.setup({
      spfxContext: this.context
    });

    return new Promise<IPropertyPaneDropdownOption[]>((resolve: (options: IPropertyPaneDropdownOption[]) => void, reject: (error: any) => void) => {

      sp.web.lists.getByTitle('Site Pages').items.top(5000).filter("OData__TopicHeader ne null and PromotedState eq '2'").get().then(data => {

        var items: IPropertyPaneDropdownOption[] = [{ key: "*", text: "No Filter" }];
        for (var k in data) {

          items.push({ key: data[k].OData__TopicHeader, text: data[k].OData__TopicHeader });
        }

        setTimeout((): void => {
          let newArray = [];
          let uniqueObject = {};
          for (let i in items) {

            // Extract the title
            let objTitle = items[i]['text'];

            // Use the title as the index
            uniqueObject[objTitle] = items[i];
          }

          // Loop to push unique object into array
          for (const i in uniqueObject) {
            newArray.push(uniqueObject[i]);
          }

          // Display the unique objects
 

          resolve(newArray);
        }, 1000);
      });
    });
  }
  protected onPropertyPaneConfigurationStart(): void {




    this.listsDropdownDisabled = !this.lists;

    if (this.lists) {
      return;
    }


    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.render();
      });


  }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let templateProperty: any;
    if (this.properties.type == "Tile" || this.properties.type == "Slide") {
      templateProperty =
        PropertyPaneChoiceGroup('layout', {

          options: [
            {
              key: 'uk-width-1-1@m', text: '1 Column',
              imageSrc: 'https://cdn0.iconfinder.com/data/icons/software-16/20/software-512.png',
              imageSize: { width: 32, height: 32 },
              selectedImageSrc: 'https://cdn0.iconfinder.com/data/icons/software-16/20/software-512.png'
            },
            {
              key: 'uk-width-1-2@m', text: '2 Column',
              imageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_2columns-512.png',
              imageSize: { width: 32, height: 32 },
              selectedImageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_2columns-512.png'
            },
            {
              key: 'uk-width-1-3@m', text: '3 Column',
              imageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_3columns-512.png',
              imageSize: { width: 32, height: 32 },
              selectedImageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_3columns-512.png', checked: true
            },
            {
              key: 'uk-width-1-4@m', text: '4 Column',
              imageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_4columns-512.png',
              imageSize: { width: 32, height: 32 },
              selectedImageSrc: 'https://cdn4.iconfinder.com/data/icons/line-icons-12/64/software_layout_4columns-512.png'
            }

          ]
        });
    } else {
      templateProperty = "";
    }




    return {
      pages: [
        {
          header: {
            description: 'Title and Layout'
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Title"
                }),

              ]
            },
            {
              groupName: "Layout",


              groupFields: [

                PropertyPaneChoiceGroup('type', {

                  options: [
                    {
                      key: 'Tile', text: 'Image on top',
                      imageSrc: 'https://cdn0.iconfinder.com/data/icons/view-1/20/vertical_slider_4-512.png',
                      imageSize: { width: 48, height: 48 },
                      selectedImageSrc: 'https://cdn0.iconfinder.com/data/icons/view-1/20/vertical_slider_4-512.png'
                    },
                    {
                      key: 'Left', text: 'Image Side',
                      imageSrc: 'https://cdn2.iconfinder.com/data/icons/interface-12/24/interface-44-512.png',
                      imageSize: { width: 48, height: 48 },
                      selectedImageSrc: 'https://cdn2.iconfinder.com/data/icons/interface-12/24/interface-44-512.png'
                    },
                    {
                      key: 'List', text: 'List with count',
                      imageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Numbered_List-256.png',
                      imageSize: { width: 48, height: 48 },
                      selectedImageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Numbered_List-256.png'
                    },
                    {
                      key: 'Slide', text: 'Carousel',
                      imageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Vertical_Align_center-256.png',
                      imageSize: { width: 48, height: 48 },
                      selectedImageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Vertical_Align_center-256.png'
                    },
                    {
                      key: 'ImageSlide', text: 'Large Carousel',
                      imageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Vertical_Distribute_Top-256.png',
                      imageSize: { width: 48, height: 48 },
                      selectedImageSrc: 'https://cdn0.iconfinder.com/data/icons/ikigai-text-and-editorial-line/32/text_Vertical_Distribute_Top-256.png'
                    }

                  ]
                }),


              ]
            }, {
              groupName: " ",


              groupFields: [

                templateProperty,


              ]
            }

          ]
        },
        {
          header: {
            description: 'Content to show'

          },
          groups: [

            {
              groupName: "",
              groupFields: [
                PropertyPaneTextField('KQLQuery', {
                  label: "Status", value: "Live"

                }),

                PropertyFieldNumber("numberValue", {
                  key: "numberValue",
                  label: "Number of results to show",
                  description: "Number of results to show",
                  value: 10,
                  maxValue: 50,

                  minValue: 1,
                  disabled: false,
                }),
                PropertyPaneCheckbox('promoted', { text: 'Show promoted', checked: true, })
              ]
            }
          ]
        },
        {
          header: {
            description: 'Sorting and Filtering'

          },
          groups: [

            {
              groupName: "",
              groupFields: [
                PropertyPaneDropdown('poll', {
                  label: "Topic Header",
                  options: this.lists,
                  disabled: this.listsDropdownDisabled,
                  selectedKey: 'Work Happy'

                }),
                PropertyPaneDropdown('sort', {
                  label: "Order by",
                  options: [
                    { key: 'PublishDate', text: 'Published Date' },
                    { key: 'Created', text: 'Last Modified Date' }


                  ],
                  selectedKey: '1',
                  disabled: this.listsDropdownDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
