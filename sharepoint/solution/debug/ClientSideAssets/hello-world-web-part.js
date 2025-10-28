(()=>{ var __RUSHSTACK_CURRENT_SCRIPT__ = document.currentScript; define("72d496c1-9c6d-4880-85c1-8fd52ef320fb_0.0.1", ["@microsoft/sp-core-library","@microsoft/sp-property-pane","@microsoft/sp-webpart-base","@microsoft/sp-lodash-subset","HelloWorldWebPartStrings"], (__WEBPACK_EXTERNAL_MODULE__676__, __WEBPACK_EXTERNAL_MODULE__877__, __WEBPACK_EXTERNAL_MODULE__642__, __WEBPACK_EXTERNAL_MODULE__529__, __WEBPACK_EXTERNAL_MODULE__275__) => { return /******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 133:
/*!**************************************************************!*\
  !*** ./lib/webparts/helloWorld/HelloWorldWebPart.module.css ***!
  \**************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! ../../../node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js */ 323);
// Imports


_node_modules_microsoft_sp_css_loader_node_modules_microsoft_load_themed_styles_lib_es6_index_js__WEBPACK_IMPORTED_MODULE_0__.loadStyles(".helloWorld_bb0f0f86{background:linear-gradient(135deg,#f5f7fa,#c3cfe2);color:\"[theme:bodyText, default: #323130]\";color:var(--bodyText);min-height:400px;overflow:hidden;padding:1.5em}.helloWorld_bb0f0f86.teams_bb0f0f86{font-family:Segoe UI,-apple-system,BlinkMacSystemFont,Roboto,Helvetica Neue,sans-serif}.headerCard_bb0f0f86{border-left:5px solid #0078d4;padding:2em}.headerCard_bb0f0f86,.teamsCard_bb0f0f86{background:#fff;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,.1);margin-bottom:1.5em}.teamsCard_bb0f0f86{border-left:5px solid #6264a7;padding:1.5em}.teamsCard_bb0f0f86 h3{color:#6264a7;margin-bottom:1em}.interactiveCard_bb0f0f86{background:#fff;border-left:5px solid #16c60c;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,.1);margin-bottom:1.5em;padding:1.5em}.interactiveCard_bb0f0f86 h3{color:#16c60c;margin-bottom:1em}.contentCard_bb0f0f86{background:#fff;border-left:5px solid #605e5c;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,.1);padding:1.5em}.contentCard_bb0f0f86 h3{color:#605e5c;margin-bottom:1em}.welcome_bb0f0f86{text-align:center}.welcome_bb0f0f86 h2{color:#0078d4;font-size:1.8em;margin:1em 0 .5em}.welcomeImage_bb0f0f86{border-radius:8px;max-width:300px;width:100%}.environmentInfo_bb0f0f86{background:#f3f9fd;border-radius:6px;color:#106ebe;font-weight:500;margin:1em 0;padding:.8em}.propertyInfo_bb0f0f86{background:#fff4ce;border-radius:6px;color:#8a6914;margin:1em 0;padding:.8em}.counterSection_bb0f0f86{text-align:center}.counterSection_bb0f0f86 p{font-size:1.2em;margin-bottom:1em}.counterDisplay_bb0f0f86{background:#16c60c;border-radius:20px;color:#fff;font-size:1.1em;font-weight:700;padding:.3em .8em}.primaryButton_bb0f0f86{background:#0078d4;border:none;border-radius:6px;color:#fff;cursor:pointer;font-weight:500;margin:0 .5em;padding:.8em 1.5em;transition:all .2s ease}.primaryButton_bb0f0f86:hover{background:#106ebe;box-shadow:0 4px 8px rgba(0,120,212,.3);transform:translateY(-1px)}.primaryButton_bb0f0f86:active{transform:translateY(0)}.secondaryButton_bb0f0f86{background:#f3f2f1;border:1px solid #d2d0ce;border-radius:6px;color:#323130;cursor:pointer;font-weight:500;margin:0 .5em;padding:.8em 1.5em;transition:all .2s ease}.secondaryButton_bb0f0f86:hover{background:#edebe9;border-color:#c8c6c4;box-shadow:0 4px 8px rgba(0,0,0,.1);transform:translateY(-1px)}.secondaryButton_bb0f0f86:active{transform:translateY(0)}.description_bb0f0f86{color:#605e5c;font-size:1.05em;line-height:1.6;margin-bottom:1.5em}.links_bb0f0f86{margin-top:1em}.links_bb0f0f86 li{margin-bottom:.5em}.links_bb0f0f86 li:hover{transform:translateX(5px);transition:transform .2s ease}.links_bb0f0f86 a{color:\"[theme:link, default:#03787c]\";color:var(--link);display:inline-block;font-weight:500;padding:.3em 0;text-decoration:none}.links_bb0f0f86 a:hover{color:\"[theme:linkHovered, default: #014446]\";color:var(--linkHovered);text-decoration:underline}.teamsSection_bb0f0f86{min-height:200px}.loadingState_bb0f0f86{color:#605e5c;padding:2em;text-align:center}.loadingState_bb0f0f86 .spinner_bb0f0f86{animation:spin_bb0f0f86 1s linear infinite;border:3px solid #f3f2f1;border-radius:50%;border-top-color:#6264a7;height:32px;margin:0 auto 1em;width:32px}@keyframes spin_bb0f0f86{0%{transform:rotate(0)}to{transform:rotate(1turn)}}.errorState_bb0f0f86{padding:2em;text-align:center}.errorState_bb0f0f86 .errorMessage_bb0f0f86{background:#fdf3f4;border-left:4px solid #d13438;border-radius:6px;color:#d13438;margin-bottom:1em;padding:1em}.emptyState_bb0f0f86{padding:2em;text-align:center}.channelsList_bb0f0f86{max-height:400px;overflow-y:auto}.teamGroup_bb0f0f86{margin-bottom:1.5em}.teamName_bb0f0f86{border-bottom:1px solid #e1dfdd;color:#6264a7;font-size:1.1em;margin:0 0 .8em;padding-bottom:.5em}.channelsGroup_bb0f0f86{margin-left:1em}.channelContainer_bb0f0f86{margin-bottom:.3em}.channelItem_bb0f0f86{align-items:center;border-radius:6px;cursor:pointer;display:flex;padding:.6em;transition:background-color .2s ease}.channelItem_bb0f0f86:hover{background:#f8f9fa}.channelItem_bb0f0f86 input[type=checkbox]{accent-color:#6264a7;margin-right:.8em;transform:scale(1.2)}.channelItem_bb0f0f86 .channelName_bb0f0f86{color:#323130;font-weight:500;margin-right:.8em}.channelItem_bb0f0f86 .channelDesc_bb0f0f86{color:#605e5c;font-size:.9em;font-style:italic;margin-right:.8em}.showFilesBtn_bb0f0f86{cursor:pointer;font-size:1.1em;margin-left:auto;padding:.2em .5em}.showFilesBtn_bb0f0f86:hover{background:#e1dfdd;border-radius:4px}.filesList_bb0f0f86{background:#f8f9fa;border-left:3px solid #6264a7;border-radius:4px;margin-left:2em;margin-top:.5em;padding:.5em}.filesUL_bb0f0f86{list-style:none;margin:0;padding:0}.fileItem_bb0f0f86,.folderItem_bb0f0f86{align-items:center;color:#323130;display:flex;font-size:.9em;padding:.3em 0}.fileItem_bb0f0f86 input[type=checkbox],.folderItem_bb0f0f86 input[type=checkbox]{accent-color:#6264a7;margin-right:.5em;transform:scale(1.1)}.folderName_bb0f0f86{border-radius:3px;cursor:pointer;padding:.2em .4em}.folderName_bb0f0f86:hover{background:#e1dfdd}.subFolderContent_bb0f0f86{background:#f0f0f0;border-left:2px solid #6264a7;border-radius:4px;margin-left:1.5em;margin-top:.5em;padding:.5em}.subFilesContainer_bb0f0f86{margin:0;padding:0}.subFileItem_bb0f0f86{align-items:center;border-bottom:1px solid #e8e8e8;color:#323130;display:flex;font-size:.85em;padding:.4em 0}.subFileItem_bb0f0f86:last-child{border-bottom:none}.subFileItem_bb0f0f86 input[type=checkbox]{accent-color:#6264a7;margin-right:.5em;transform:scale(1.1)}.subFileItem_bb0f0f86 span{word-break:break-all}.teamsHeader_bb0f0f86{align-items:center;display:flex;justify-content:space-between;margin-bottom:1em}.teamsHeader_bb0f0f86 h3{margin:0}.viewToggle_bb0f0f86{display:flex;gap:.5em}.toggleButton_bb0f0f86{background:#f3f2f1;border:1px solid #d2d0ce;border-radius:4px;color:#323130;cursor:pointer;font-size:.85em;padding:.5em 1em;transition:all .2s ease}.toggleButton_bb0f0f86:hover{background:#edebe9;border-color:#c8c6c4}.toggleButton_bb0f0f86.active_bb0f0f86{background:#6264a7;border-color:#6264a7;color:#fff}.toggleButton_bb0f0f86.active_bb0f0f86:hover{background:#5a5da6}.treeContainer_bb0f0f86{background:#fff;border:1px solid #e1dfdd;border-radius:6px;max-height:500px;overflow-y:auto}.treeItem_bb0f0f86{border-bottom:1px solid #f8f9fa}.treeItem_bb0f0f86:hover{background:#f8f9fa}.treeItem_bb0f0f86:last-child{border-bottom:none}.treeItemContent_bb0f0f86{align-items:center;display:flex;min-height:32px;padding:.4em .6em}.expandIcon_bb0f0f86{color:#605e5c;cursor:pointer;font-size:.8em;margin-right:.5em;text-align:center;-webkit-user-select:none;-ms-user-select:none;user-select:none;width:16px}.expandIcon_bb0f0f86:hover{color:#323130}.itemIcon_bb0f0f86{font-size:1em;margin-right:.5em;min-width:20px;text-align:center}.treeItemLabel_bb0f0f86{align-items:center;cursor:pointer;display:flex;flex:1}.treeItemLabel_bb0f0f86 input[type=checkbox]{accent-color:#6264a7;margin-right:.6em;transform:scale(1.1)}.treeItemLabel_bb0f0f86 .itemName_bb0f0f86{color:#323130;font-weight:500;margin-right:.6em}.treeItemLabel_bb0f0f86 .itemDesc_bb0f0f86{color:#605e5c;font-size:.85em;font-style:italic;margin-right:.6em}.treeItemLabel_bb0f0f86 .loadingIcon_bb0f0f86{animation:spin_bb0f0f86 1s linear infinite;color:#6264a7;font-size:.9em}\n/*# sourceMappingURL=data:application/json;base64,eyJ2ZXJzaW9uIjozLCJzb3VyY2VzIjpbImZpbGU6Ly8vQzovVXNlcnMvUmFqZXNoLmFsZGEvRG93bmxvYWRzL1NwJTIwdGVzdC9zcmMvd2VicGFydHMvaGVsbG9Xb3JsZC9IZWxsb1dvcmxkV2ViUGFydC5tb2R1bGUuc2NzcyJdLCJuYW1lcyI6W10sIm1hcHBpbmdzIjoiQUFFQSxxQkFLRSxrREFBQSxDQUZBLDBDQUFBLENBQ0EscUJBQUEsQ0FFQSxnQkFBQSxDQUxBLGVBQUEsQ0FDQSxhQUlBLENBRUEsb0NBQ0Usc0ZBQUEsQ0FJSixxQkFNRSw2QkFBQSxDQUhBLFdBR0EsQ0FHRix5Q0FSRSxlQUFBLENBQ0Esa0JBQUEsQ0FHQSxvQ0FBQSxDQURBLG1CQVdBLENBTkYsb0JBTUUsNkJBQUEsQ0FIQSxhQUdBLENBRUEsdUJBQ0UsYUFBQSxDQUNBLGlCQUFBLENBSUosMEJBQ0UsZUFBQSxDQUtBLDZCQUFBLENBSkEsa0JBQUEsQ0FHQSxvQ0FBQSxDQURBLG1CQUFBLENBREEsYUFHQSxDQUVBLDZCQUNFLGFBQUEsQ0FDQSxpQkFBQSxDQUlKLHNCQUNFLGVBQUEsQ0FJQSw2QkFBQSxDQUhBLGtCQUFBLENBRUEsb0NBQUEsQ0FEQSxhQUVBLENBRUEseUJBQ0UsYUFBQSxDQUNBLGlCQUFBLENBSUosa0JBQ0UsaUJBQUEsQ0FFQSxxQkFDRSxhQUFBLENBRUEsZUFBQSxDQURBLGlCQUNBLENBSUosdUJBR0UsaUJBQUEsQ0FEQSxlQUFBLENBREEsVUFFQSxDQUdGLDBCQUNFLGtCQUFBLENBRUEsaUJBQUEsQ0FFQSxhQUFBLENBQ0EsZUFBQSxDQUZBLFlBQUEsQ0FGQSxZQUlBLENBR0YsdUJBQ0Usa0JBQUEsQ0FFQSxpQkFBQSxDQUVBLGFBQUEsQ0FEQSxZQUFBLENBRkEsWUFHQSxDQUdGLHlCQUNFLGlCQUFBLENBRUEsMkJBQ0UsZUFBQSxDQUNBLGlCQUFBLENBSUoseUJBQ0Usa0JBQUEsQ0FHQSxrQkFBQSxDQUZBLFVBQUEsQ0FJQSxlQUFBLENBREEsZUFBQSxDQUZBLGlCQUdBLENBR0Ysd0JBQ0Usa0JBQUEsQ0FFQSxXQUFBLENBRUEsaUJBQUEsQ0FIQSxVQUFBLENBS0EsY0FBQSxDQUNBLGVBQUEsQ0FGQSxhQUFBLENBRkEsa0JBQUEsQ0FLQSx1QkFBQSxDQUVBLDhCQUNFLGtCQUFBLENBRUEsdUNBQUEsQ0FEQSwwQkFDQSxDQUdGLCtCQUNFLHVCQUFBLENBSUosMEJBQ0Usa0JBQUEsQ0FFQSx3QkFBQSxDQUVBLGlCQUFBLENBSEEsYUFBQSxDQUtBLGNBQUEsQ0FDQSxlQUFBLENBRkEsYUFBQSxDQUZBLGtCQUFBLENBS0EsdUJBQUEsQ0FFQSxnQ0FDRSxrQkFBQSxDQUNBLG9CQUFBLENBRUEsbUNBQUEsQ0FEQSwwQkFDQSxDQUdGLGlDQUNFLHVCQUFBLENBSUosc0JBR0UsYUFBQSxDQUZBLGdCQUFBLENBQ0EsZUFBQSxDQUVBLG1CQUFBLENBR0YsZ0JBQ0UsY0FBQSxDQUVBLG1CQUNFLGtCQUFBLENBRUEseUJBQ0UseUJBQUEsQ0FDQSw2QkFBQSxDQUlKLGtCQUVFLHFDQUFBLENBQ0EsaUJBQUEsQ0FFQSxvQkFBQSxDQURBLGVBQUEsQ0FFQSxjQUFBLENBTEEsb0JBS0EsQ0FFQSx3QkFFRSw2Q0FBQSxDQUNBLHdCQUFBLENBRkEseUJBRUEsQ0FNTix1QkFDRSxnQkFBQSxDQUdGLHVCQUdFLGFBQUEsQ0FEQSxXQUFBLENBREEsaUJBRUEsQ0FFQSx5Q0FNRSwwQ0FBQSxDQUZBLHdCQUFBLENBQ0EsaUJBQUEsQ0FEQSx3QkFBQSxDQUZBLFdBQUEsQ0FLQSxpQkFBQSxDQU5BLFVBTUEsQ0FJSix5QkFDRSxHQUFLLG1CQUFBLENBQ0wsR0FBTyx1QkFBQSxDQUFBLENBR1QscUJBRUUsV0FBQSxDQURBLGlCQUNBLENBRUEsNENBSUUsa0JBQUEsQ0FFQSw2QkFBQSxDQURBLGlCQUFBLENBSkEsYUFBQSxDQUNBLGlCQUFBLENBQ0EsV0FHQSxDQUlKLHFCQUVFLFdBQUEsQ0FEQSxpQkFDQSxDQUdGLHVCQUNFLGdCQUFBLENBQ0EsZUFBQSxDQUdGLG9CQUNFLG1CQUFBLENBR0YsbUJBSUUsK0JBQUEsQ0FIQSxhQUFBLENBSUEsZUFBQSxDQUhBLGVBQUEsQ0FDQSxtQkFFQSxDQUdGLHdCQUNFLGVBQUEsQ0FHRiwyQkFDRSxrQkFBQSxDQUdGLHNCQUVFLGtCQUFBLENBRUEsaUJBQUEsQ0FDQSxjQUFBLENBSkEsWUFBQSxDQUVBLFlBQUEsQ0FHQSxvQ0FBQSxDQUVBLDRCQUNFLGtCQUFBLENBR0YsMkNBR0Usb0JBQUEsQ0FGQSxpQkFBQSxDQUNBLG9CQUNBLENBR0YsNENBRUUsYUFBQSxDQURBLGVBQUEsQ0FFQSxpQkFBQSxDQUdGLDRDQUNFLGFBQUEsQ0FDQSxjQUFBLENBQ0EsaUJBQUEsQ0FDQSxpQkFBQSxDQUlKLHVCQUVFLGNBQUEsQ0FFQSxlQUFBLENBSEEsZ0JBQUEsQ0FFQSxpQkFDQSxDQUVBLDZCQUNFLGtCQUFBLENBQ0EsaUJBQUEsQ0FJSixvQkFJRSxrQkFBQSxDQUVBLDZCQUFBLENBREEsaUJBQUEsQ0FKQSxlQUFBLENBQ0EsZUFBQSxDQUNBLFlBR0EsQ0FHRixrQkFHRSxlQUFBLENBRkEsUUFBQSxDQUNBLFNBQ0EsQ0FHRix3Q0FFRSxrQkFBQSxDQUVBLGFBQUEsQ0FIQSxZQUFBLENBSUEsY0FBQSxDQUZBLGNBRUEsQ0FFQSxrRkFHRSxvQkFBQSxDQUZBLGlCQUFBLENBQ0Esb0JBQ0EsQ0FJSixxQkFHRSxpQkFBQSxDQUZBLGNBQUEsQ0FDQSxpQkFDQSxDQUVBLDJCQUNFLGtCQUFBLENBSUosMkJBSUUsa0JBQUEsQ0FFQSw2QkFBQSxDQURBLGlCQUFBLENBSkEsaUJBQUEsQ0FDQSxlQUFBLENBQ0EsWUFHQSxDQUdGLDRCQUNFLFFBQUEsQ0FDQSxTQUFBLENBR0Ysc0JBRUUsa0JBQUEsQ0FJQSwrQkFBQSxDQUZBLGFBQUEsQ0FIQSxZQUFBLENBSUEsZUFBQSxDQUZBLGNBR0EsQ0FFQSxpQ0FDRSxrQkFBQSxDQUdGLDJDQUdFLG9CQUFBLENBRkEsaUJBQUEsQ0FDQSxvQkFDQSxDQUdGLDJCQUNFLG9CQUFBLENBS0osc0JBR0Usa0JBQUEsQ0FGQSxZQUFBLENBQ0EsNkJBQUEsQ0FFQSxpQkFBQSxDQUVBLHlCQUNFLFFBQUEsQ0FJSixxQkFDRSxZQUFBLENBQ0EsUUFBQSxDQUdGLHVCQUNFLGtCQUFBLENBRUEsd0JBQUEsQ0FFQSxpQkFBQSxDQUhBLGFBQUEsQ0FJQSxjQUFBLENBQ0EsZUFBQSxDQUhBLGdCQUFBLENBSUEsdUJBQUEsQ0FFQSw2QkFDRSxrQkFBQSxDQUNBLG9CQUFBLENBR0YsdUNBQ0Usa0JBQUEsQ0FFQSxvQkFBQSxDQURBLFVBQ0EsQ0FFQSw2Q0FDRSxrQkFBQSxDQUtOLHdCQUtFLGVBQUEsQ0FGQSx3QkFBQSxDQUNBLGlCQUFBLENBSEEsZ0JBQUEsQ0FDQSxlQUdBLENBR0YsbUJBQ0UsK0JBQUEsQ0FFQSx5QkFDRSxrQkFBQSxDQUdGLDhCQUNFLGtCQUFBLENBSUosMEJBRUUsa0JBQUEsQ0FEQSxZQUFBLENBR0EsZUFBQSxDQURBLGlCQUNBLENBR0YscUJBS0UsYUFBQSxDQUZBLGNBQUEsQ0FHQSxjQUFBLENBQ0EsaUJBQUEsQ0FMQSxpQkFBQSxDQUVBLHdCQUFBLENBQUEsb0JBQUEsQ0FBQSxnQkFBQSxDQUhBLFVBTUEsQ0FFQSwyQkFDRSxhQUFBLENBSUosbUJBRUUsYUFBQSxDQURBLGlCQUFBLENBRUEsY0FBQSxDQUNBLGlCQUFBLENBR0Ysd0JBRUUsa0JBQUEsQ0FDQSxjQUFBLENBRkEsWUFBQSxDQUdBLE1BQUEsQ0FFQSw2Q0FHRSxvQkFBQSxDQUZBLGlCQUFBLENBQ0Esb0JBQ0EsQ0FHRiwyQ0FFRSxhQUFBLENBREEsZUFBQSxDQUVBLGlCQUFBLENBR0YsMkNBQ0UsYUFBQSxDQUNBLGVBQUEsQ0FDQSxpQkFBQSxDQUNBLGlCQUFBLENBR0YsOENBRUUsMENBQUEsQ0FEQSxhQUFBLENBRUEsY0FBQSIsImZpbGUiOiJIZWxsb1dvcmxkV2ViUGFydC5tb2R1bGUuY3NzIn0= */", true);

// Exports
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = ({
  helloWorld_bb0f0f86: "helloWorld_bb0f0f86",
  teams_bb0f0f86: "teams_bb0f0f86",
  headerCard_bb0f0f86: "headerCard_bb0f0f86",
  teamsCard_bb0f0f86: "teamsCard_bb0f0f86",
  interactiveCard_bb0f0f86: "interactiveCard_bb0f0f86",
  contentCard_bb0f0f86: "contentCard_bb0f0f86",
  welcome_bb0f0f86: "welcome_bb0f0f86",
  welcomeImage_bb0f0f86: "welcomeImage_bb0f0f86",
  environmentInfo_bb0f0f86: "environmentInfo_bb0f0f86",
  propertyInfo_bb0f0f86: "propertyInfo_bb0f0f86",
  counterSection_bb0f0f86: "counterSection_bb0f0f86",
  counterDisplay_bb0f0f86: "counterDisplay_bb0f0f86",
  primaryButton_bb0f0f86: "primaryButton_bb0f0f86",
  secondaryButton_bb0f0f86: "secondaryButton_bb0f0f86",
  description_bb0f0f86: "description_bb0f0f86",
  links_bb0f0f86: "links_bb0f0f86",
  teamsSection_bb0f0f86: "teamsSection_bb0f0f86",
  loadingState_bb0f0f86: "loadingState_bb0f0f86",
  spinner_bb0f0f86: "spinner_bb0f0f86",
  spin_bb0f0f86: "spin_bb0f0f86",
  errorState_bb0f0f86: "errorState_bb0f0f86",
  errorMessage_bb0f0f86: "errorMessage_bb0f0f86",
  emptyState_bb0f0f86: "emptyState_bb0f0f86",
  channelsList_bb0f0f86: "channelsList_bb0f0f86",
  teamGroup_bb0f0f86: "teamGroup_bb0f0f86",
  teamName_bb0f0f86: "teamName_bb0f0f86",
  channelsGroup_bb0f0f86: "channelsGroup_bb0f0f86",
  channelContainer_bb0f0f86: "channelContainer_bb0f0f86",
  channelItem_bb0f0f86: "channelItem_bb0f0f86",
  channelName_bb0f0f86: "channelName_bb0f0f86",
  channelDesc_bb0f0f86: "channelDesc_bb0f0f86",
  showFilesBtn_bb0f0f86: "showFilesBtn_bb0f0f86",
  filesList_bb0f0f86: "filesList_bb0f0f86",
  filesUL_bb0f0f86: "filesUL_bb0f0f86",
  fileItem_bb0f0f86: "fileItem_bb0f0f86",
  folderItem_bb0f0f86: "folderItem_bb0f0f86",
  folderName_bb0f0f86: "folderName_bb0f0f86",
  subFolderContent_bb0f0f86: "subFolderContent_bb0f0f86",
  subFilesContainer_bb0f0f86: "subFilesContainer_bb0f0f86",
  subFileItem_bb0f0f86: "subFileItem_bb0f0f86",
  teamsHeader_bb0f0f86: "teamsHeader_bb0f0f86",
  viewToggle_bb0f0f86: "viewToggle_bb0f0f86",
  toggleButton_bb0f0f86: "toggleButton_bb0f0f86",
  active_bb0f0f86: "active_bb0f0f86",
  treeContainer_bb0f0f86: "treeContainer_bb0f0f86",
  treeItem_bb0f0f86: "treeItem_bb0f0f86",
  treeItemContent_bb0f0f86: "treeItemContent_bb0f0f86",
  expandIcon_bb0f0f86: "expandIcon_bb0f0f86",
  itemIcon_bb0f0f86: "itemIcon_bb0f0f86",
  treeItemLabel_bb0f0f86: "treeItemLabel_bb0f0f86",
  itemName_bb0f0f86: "itemName_bb0f0f86",
  itemDesc_bb0f0f86: "itemDesc_bb0f0f86",
  loadingIcon_bb0f0f86: "loadingIcon_bb0f0f86"
});


/***/ }),

/***/ 1:
/*!******************************************************************!*\
  !*** ./lib/webparts/helloWorld/HelloWorldWebPart.module.scss.js ***!
  \******************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
__webpack_require__(/*! ./HelloWorldWebPart.module.css */ 133);
var styles = {
    helloWorld: 'helloWorld_bb0f0f86',
    teams: 'teams_bb0f0f86',
    headerCard: 'headerCard_bb0f0f86',
    teamsCard: 'teamsCard_bb0f0f86',
    interactiveCard: 'interactiveCard_bb0f0f86',
    contentCard: 'contentCard_bb0f0f86',
    welcome: 'welcome_bb0f0f86',
    welcomeImage: 'welcomeImage_bb0f0f86',
    environmentInfo: 'environmentInfo_bb0f0f86',
    propertyInfo: 'propertyInfo_bb0f0f86',
    counterSection: 'counterSection_bb0f0f86',
    counterDisplay: 'counterDisplay_bb0f0f86',
    primaryButton: 'primaryButton_bb0f0f86',
    secondaryButton: 'secondaryButton_bb0f0f86',
    description: 'description_bb0f0f86',
    links: 'links_bb0f0f86',
    teamsSection: 'teamsSection_bb0f0f86',
    loadingState: 'loadingState_bb0f0f86',
    spinner: 'spinner_bb0f0f86',
    spin: 'spin_bb0f0f86',
    errorState: 'errorState_bb0f0f86',
    errorMessage: 'errorMessage_bb0f0f86',
    emptyState: 'emptyState_bb0f0f86',
    channelsList: 'channelsList_bb0f0f86',
    teamGroup: 'teamGroup_bb0f0f86',
    teamName: 'teamName_bb0f0f86',
    channelsGroup: 'channelsGroup_bb0f0f86',
    channelContainer: 'channelContainer_bb0f0f86',
    channelItem: 'channelItem_bb0f0f86',
    channelName: 'channelName_bb0f0f86',
    channelDesc: 'channelDesc_bb0f0f86',
    showFilesBtn: 'showFilesBtn_bb0f0f86',
    filesList: 'filesList_bb0f0f86',
    filesUL: 'filesUL_bb0f0f86',
    fileItem: 'fileItem_bb0f0f86',
    folderItem: 'folderItem_bb0f0f86',
    folderName: 'folderName_bb0f0f86',
    subFolderContent: 'subFolderContent_bb0f0f86',
    subFilesContainer: 'subFilesContainer_bb0f0f86',
    subFileItem: 'subFileItem_bb0f0f86',
    teamsHeader: 'teamsHeader_bb0f0f86',
    viewToggle: 'viewToggle_bb0f0f86',
    toggleButton: 'toggleButton_bb0f0f86',
    active: 'active_bb0f0f86',
    treeContainer: 'treeContainer_bb0f0f86',
    treeItem: 'treeItem_bb0f0f86',
    treeItemContent: 'treeItemContent_bb0f0f86',
    expandIcon: 'expandIcon_bb0f0f86',
    itemIcon: 'itemIcon_bb0f0f86',
    treeItemLabel: 'treeItemLabel_bb0f0f86',
    itemName: 'itemName_bb0f0f86',
    itemDesc: 'itemDesc_bb0f0f86',
    loadingIcon: 'loadingIcon_bb0f0f86'
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (styles);


/***/ }),

/***/ 266:
/*!**********************************************************!*\
  !*** ./lib/webparts/helloWorld/services/TeamsService.js ***!
  \**********************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   TeamsService: () => (/* binding */ TeamsService)
/* harmony export */ });
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
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
var TeamsService = /** @class */ (function () {
    function TeamsService(graphClient) {
        this.graphClient = graphClient;
    }
    TeamsService.prototype.getUserTeams = function () {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api('/me/joinedTeams')
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            throw new Error('No teams data received');
                        }
                        return [2 /*return*/, response.value.map(function (team) { return ({
                                id: team.id,
                                displayName: team.displayName,
                                description: team.description
                            }); })];
                    case 2:
                        error_1 = _a.sent();
                        throw this._handleError(error_1);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getTeamChannels = function (teamId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            throw new Error('No channels data received');
                        }
                        return [2 /*return*/, response.value.map(function (channel) { return ({
                                id: channel.id,
                                displayName: channel.displayName,
                                description: channel.description
                            }); })];
                    case 2:
                        error_2 = _a.sent();
                        throw this._handleError(error_2);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getAllChannelsForUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teams, error, allChannels, _loop_1, this_1, _i, teams_1, team, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 6, , 7]);
                        return [4 /*yield*/, this.getUserTeams()];
                    case 1:
                        teams = _a.sent();
                        if (teams.length === 0) {
                            error = {
                                type: 'NO_TEAMS',
                                message: 'User is not a member of any teams'
                            };
                            throw error;
                        }
                        allChannels = [];
                        _loop_1 = function (team) {
                            var channels, channelsWithTeamInfo, error_4;
                            return __generator(this, function (_b) {
                                switch (_b.label) {
                                    case 0:
                                        _b.trys.push([0, 2, , 3]);
                                        return [4 /*yield*/, this_1.getTeamChannels(team.id)];
                                    case 1:
                                        channels = _b.sent();
                                        channelsWithTeamInfo = channels.map(function (channel) { return (__assign(__assign({}, channel), { teamId: team.id, teamName: team.displayName })); });
                                        allChannels.push.apply(allChannels, channelsWithTeamInfo);
                                        return [3 /*break*/, 3];
                                    case 2:
                                        error_4 = _b.sent();
                                        console.warn("Failed to get channels for team ".concat(team.displayName, ":"), error_4);
                                        return [3 /*break*/, 3];
                                    case 3: return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _i = 0, teams_1 = teams;
                        _a.label = 2;
                    case 2:
                        if (!(_i < teams_1.length)) return [3 /*break*/, 5];
                        team = teams_1[_i];
                        return [5 /*yield**/, _loop_1(team)];
                    case 3:
                        _a.sent();
                        _a.label = 4;
                    case 4:
                        _i++;
                        return [3 /*break*/, 2];
                    case 5: return [2 /*return*/, allChannels];
                    case 6:
                        error_3 = _a.sent();
                        if (error_3.type) {
                            throw error_3;
                        }
                        throw this._handleError(error_3);
                    case 7: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getChannelFilesFolder = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var response, error_5;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels/").concat(channelId, "/filesFolder"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        console.log('Raw filesFolder response:', JSON.stringify(response, null, 2));
                        return [2 /*return*/, {
                                id: response.id,
                                name: response.name,
                                webUrl: response.webUrl,
                                size: response.size,
                                folder: response.folder,
                                createdDateTime: response.createdDateTime,
                                lastModifiedDateTime: response.lastModifiedDateTime
                            }];
                    case 2:
                        error_5 = _a.sent();
                        throw this._handleError(error_5);
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.getDriveItemChildren = function (driveId, itemId) {
        if (itemId === void 0) { itemId = 'root'; }
        return __awaiter(this, void 0, void 0, function () {
            var response, error_6;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this.graphClient
                                .api("/drives/".concat(driveId, "/items/").concat(itemId, "/children"))
                                .get()];
                    case 1:
                        response = _a.sent();
                        if (!response || !response.value) {
                            return [2 /*return*/, []];
                        }
                        return [2 /*return*/, response.value.map(function (item) { return ({
                                id: item.id,
                                name: item.name,
                                webUrl: item.webUrl,
                                size: item.size,
                                folder: item.folder,
                                file: item.file,
                                createdDateTime: item.createdDateTime,
                                lastModifiedDateTime: item.lastModifiedDateTime
                            }); })];
                    case 2:
                        error_6 = _a.sent();
                        console.warn('Failed to get drive item children:', error_6);
                        return [2 /*return*/, []];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.buildTeamsTreeView = function () {
        return __awaiter(this, void 0, void 0, function () {
            var teams, treeNodes, _i, teams_2, team, teamNode, channels, _a, channels_1, channel, channelNode, error_7, error_8;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 9, , 10]);
                        return [4 /*yield*/, this.getUserTeams()];
                    case 1:
                        teams = _b.sent();
                        treeNodes = [];
                        _i = 0, teams_2 = teams;
                        _b.label = 2;
                    case 2:
                        if (!(_i < teams_2.length)) return [3 /*break*/, 8];
                        team = teams_2[_i];
                        teamNode = {
                            id: team.id,
                            name: team.displayName,
                            type: 'team',
                            description: team.description,
                            isExpanded: false,
                            children: []
                        };
                        _b.label = 3;
                    case 3:
                        _b.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, this.getTeamChannels(team.id)];
                    case 4:
                        channels = _b.sent();
                        for (_a = 0, channels_1 = channels; _a < channels_1.length; _a++) {
                            channel = channels_1[_a];
                            channelNode = {
                                id: "".concat(team.id, "_").concat(channel.id),
                                name: channel.displayName,
                                type: 'channel',
                                description: channel.description,
                                parentId: team.id,
                                teamId: team.id,
                                channelId: channel.id,
                                isExpanded: false,
                                children: [],
                                childCount: 0,
                                isLoading: false
                            };
                            if (teamNode.children) {
                                teamNode.children.push(channelNode);
                            }
                        }
                        return [3 /*break*/, 6];
                    case 5:
                        error_7 = _b.sent();
                        console.warn("Failed to get channels for team ".concat(team.displayName, ":"), error_7);
                        return [3 /*break*/, 6];
                    case 6:
                        treeNodes.push(teamNode);
                        _b.label = 7;
                    case 7:
                        _i++;
                        return [3 /*break*/, 2];
                    case 8: return [2 /*return*/, treeNodes];
                    case 9:
                        error_8 = _b.sent();
                        throw this._handleError(error_8);
                    case 10: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.loadChannelFiles = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var filesFolder, driveResponse, rootItems, channelFolderItems, _i, rootItems_1, item, folderContents, folderData, folderError_1, error_9;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 13, , 14]);
                        return [4 /*yield*/, this.getChannelFilesFolder(teamId, channelId)];
                    case 1:
                        filesFolder = _a.sent();
                        console.log('FilesFolder response:', filesFolder);
                        // Extract the drive and item information from the filesFolder
                        // The filesFolder should contain drive information
                        if (!filesFolder.webUrl) {
                            console.warn('No webUrl in filesFolder response');
                            return [2 /*return*/, []];
                        }
                        return [4 /*yield*/, this.graphClient
                                .api("/groups/".concat(teamId, "/drive"))
                                .get()];
                    case 2:
                        driveResponse = _a.sent();
                        if (!driveResponse) {
                            console.warn('No drive response for team');
                            return [2 /*return*/, []];
                        }
                        console.log('Drive response:', driveResponse);
                        return [4 /*yield*/, this.getDriveItemChildren(driveResponse.id, 'root')];
                    case 3:
                        rootItems = _a.sent();
                        console.log('Root items:', rootItems);
                        channelFolderItems = [];
                        _i = 0, rootItems_1 = rootItems;
                        _a.label = 4;
                    case 4:
                        if (!(_i < rootItems_1.length)) return [3 /*break*/, 7];
                        item = rootItems_1[_i];
                        if (!item.folder) return [3 /*break*/, 6];
                        return [4 /*yield*/, this.getDriveItemChildren(driveResponse.id, item.id)];
                    case 5:
                        folderContents = _a.sent();
                        console.log("Contents of ".concat(item.name, ":"), folderContents);
                        channelFolderItems = channelFolderItems.concat(folderContents);
                        _a.label = 6;
                    case 6:
                        _i++;
                        return [3 /*break*/, 4];
                    case 7:
                        if (!(channelFolderItems.length === 0)) return [3 /*break*/, 12];
                        _a.label = 8;
                    case 8:
                        _a.trys.push([8, 11, , 12]);
                        folderData = filesFolder;
                        if (!(folderData.parentReference && folderData.parentReference.driveId)) return [3 /*break*/, 10];
                        return [4 /*yield*/, this.getDriveItemChildren(folderData.parentReference.driveId, folderData.id)];
                    case 9:
                        channelFolderItems = _a.sent();
                        _a.label = 10;
                    case 10: return [3 /*break*/, 12];
                    case 11:
                        folderError_1 = _a.sent();
                        console.warn('Fallback method failed:', folderError_1);
                        return [3 /*break*/, 12];
                    case 12: return [2 /*return*/, this._convertDriveItemsToTreeNodes(channelFolderItems, driveResponse.id, "".concat(teamId, "_").concat(channelId))];
                    case 13:
                        error_9 = _a.sent();
                        console.error("Failed to load files for channel ".concat(channelId, ":"), error_9);
                        return [2 /*return*/, []];
                    case 14: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype.loadChannelFilesDirect = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var filesFolder, fullResponse, children, error_10;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 5, , 6]);
                        console.log("\uD83D\uDD0D Starting loadChannelFilesDirect for teamId: ".concat(teamId, ", channelId: ").concat(channelId));
                        return [4 /*yield*/, this.getChannelFilesFolder(teamId, channelId)];
                    case 1:
                        filesFolder = _a.sent();
                        console.log('📁 FilesFolder data:', JSON.stringify(filesFolder, null, 2));
                        if (!filesFolder.id) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.graphClient
                                .api("/teams/".concat(teamId, "/channels/").concat(channelId, "/filesFolder"))
                                .select('id,name,parentReference,webUrl')
                                .get()];
                    case 2:
                        fullResponse = _a.sent();
                        console.log('🔧 Full filesFolder response:', JSON.stringify(fullResponse, null, 2));
                        if (!(fullResponse.parentReference && fullResponse.parentReference.driveId)) return [3 /*break*/, 4];
                        console.log("\uD83D\uDCC2 Found drive ID: ".concat(fullResponse.parentReference.driveId, ", item ID: ").concat(fullResponse.id));
                        return [4 /*yield*/, this.getDriveItemChildren(fullResponse.parentReference.driveId, fullResponse.id)];
                    case 3:
                        children = _a.sent();
                        console.log("\uD83D\uDCC4 Found ".concat(children.length, " items in channel folder"));
                        if (children.length > 0) {
                            console.log('✅ Successfully loaded files:', children.map(function (c) { return c.name; }));
                            return [2 /*return*/, this._convertDriveItemsToTreeNodes(children, fullResponse.parentReference.driveId, "".concat(teamId, "_").concat(channelId))];
                        }
                        _a.label = 4;
                    case 4:
                        console.log('⚠️ No files found or no drive reference available');
                        return [2 /*return*/, []];
                    case 5:
                        error_10 = _a.sent();
                        console.error("\u274C Direct approach failed for channel ".concat(channelId, ":"), error_10);
                        console.error('Error details:', JSON.stringify(error_10, null, 2));
                        return [2 /*return*/, []];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    TeamsService.prototype._convertDriveItemsToTreeNodes = function (driveItems, driveId, parentId) {
        return driveItems.map(function (item) {
            var _a;
            return ({
                id: "".concat(parentId, "_").concat(item.id),
                name: item.name,
                type: item.folder ? 'folder' : 'file',
                parentId: parentId,
                webUrl: item.webUrl,
                size: item.size,
                driveId: driveId,
                itemId: item.id,
                childCount: ((_a = item.folder) === null || _a === void 0 ? void 0 : _a.childCount) || 0,
                isExpanded: false,
                children: [],
                isLoading: false
            });
        });
    };
    TeamsService.prototype._handleError = function (error) {
        console.error('Teams Service Error:', error);
        if (error.code === 'Forbidden' || error.status === 403) {
            return {
                type: 'NO_PERMISSIONS',
                message: 'Insufficient permissions to access Teams data. Please contact your administrator.',
                originalError: error
            };
        }
        if (error.code === 'NetworkError' || error.name === 'NetworkError') {
            return {
                type: 'NETWORK_ERROR',
                message: 'Network error occurred while fetching Teams data. Please check your connection and try again.',
                originalError: error
            };
        }
        return {
            type: 'UNKNOWN',
            message: "An unexpected error occurred: ".concat(error.message || 'Unknown error'),
            originalError: error
        };
    };
    return TeamsService;
}());



/***/ }),

/***/ 323:
/*!***********************************************************************************************************!*\
  !*** ./node_modules/@microsoft/sp-css-loader/node_modules/@microsoft/load-themed-styles/lib-es6/index.js ***!
  \***********************************************************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   ClearStyleOptions: () => (/* binding */ ClearStyleOptions),
/* harmony export */   Mode: () => (/* binding */ Mode),
/* harmony export */   clearStyles: () => (/* binding */ clearStyles),
/* harmony export */   configureLoadStyles: () => (/* binding */ configureLoadStyles),
/* harmony export */   configureRunMode: () => (/* binding */ configureRunMode),
/* harmony export */   detokenize: () => (/* binding */ detokenize),
/* harmony export */   flush: () => (/* binding */ flush),
/* harmony export */   loadStyles: () => (/* binding */ loadStyles),
/* harmony export */   loadTheme: () => (/* binding */ loadTheme),
/* harmony export */   splitStyles: () => (/* binding */ splitStyles)
/* harmony export */ });
// Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// See LICENSE in the project root for license information.
var __assign = (undefined && undefined.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
/**
 * In sync mode, styles are registered as style elements synchronously with loadStyles() call.
 * In async mode, styles are buffered and registered as batch in async timer for performance purpose.
 */
var Mode;
(function (Mode) {
    Mode[Mode["sync"] = 0] = "sync";
    Mode[Mode["async"] = 1] = "async";
})(Mode || (Mode = {}));
/**
 * Themable styles and non-themable styles are tracked separately
 * Specify ClearStyleOptions when calling clearStyles API to specify which group of registered styles should be cleared.
 */
var ClearStyleOptions;
(function (ClearStyleOptions) {
    /** only themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyThemable"] = 1] = "onlyThemable";
    /** only non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["onlyNonThemable"] = 2] = "onlyNonThemable";
    /** both themable and non-themable styles will be cleared */
    ClearStyleOptions[ClearStyleOptions["all"] = 3] = "all";
})(ClearStyleOptions || (ClearStyleOptions = {}));
// Store the theming state in __themeState__ global scope for reuse in the case of duplicate
// load-themed-styles hosted on the page.
var _root = typeof window === 'undefined' ? __webpack_require__.g : window; // eslint-disable-line @typescript-eslint/no-explicit-any
// Nonce string to inject into script tag if one provided. This is used in CSP (Content Security Policy).
var _styleNonce = _root && _root.CSPSettings && _root.CSPSettings.nonce;
var _themeState = initializeThemeState();
/**
 * Matches theming tokens. For example, "[theme: themeSlotName, default: #FFF]" (including the quotes).
 */
var _themeTokenRegex = /[\'\"]\[theme:\s*(\w+)\s*(?:\,\s*default:\s*([\\"\']?[\.\,\(\)\#\-\s\w]*[\.\,\(\)\#\-\w][\"\']?))?\s*\][\'\"]/g;
var now = function () {
    return typeof performance !== 'undefined' && !!performance.now ? performance.now() : Date.now();
};
function measure(func) {
    var start = now();
    func();
    var end = now();
    _themeState.perf.duration += end - start;
}
/**
 * initialize global state object
 */
function initializeThemeState() {
    var state = _root.__themeState__ || {
        theme: undefined,
        lastStyleElement: undefined,
        registeredStyles: []
    };
    if (!state.runState) {
        state = __assign(__assign({}, state), { perf: {
                count: 0,
                duration: 0
            }, runState: {
                flushTimer: 0,
                mode: Mode.sync,
                buffer: []
            } });
    }
    if (!state.registeredThemableStyles) {
        state = __assign(__assign({}, state), { registeredThemableStyles: [] });
    }
    _root.__themeState__ = state;
    return state;
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load
 * event is fired.
 * @param {string | ThemableArray} styles Themable style text to register.
 * @param {boolean} loadAsync When true, always load styles in async mode, irrespective of current sync mode.
 */
function loadStyles(styles, loadAsync) {
    if (loadAsync === void 0) { loadAsync = false; }
    measure(function () {
        var styleParts = Array.isArray(styles) ? styles : splitStyles(styles);
        var _a = _themeState.runState, mode = _a.mode, buffer = _a.buffer, flushTimer = _a.flushTimer;
        if (loadAsync || mode === Mode.async) {
            buffer.push(styleParts);
            if (!flushTimer) {
                _themeState.runState.flushTimer = asyncLoadStyles();
            }
        }
        else {
            applyThemableStyles(styleParts);
        }
    });
}
/**
 * Allows for customizable loadStyles logic. e.g. for server side rendering application
 * @param {(processedStyles: string, rawStyles?: string | ThemableArray) => void}
 * a loadStyles callback that gets called when styles are loaded or reloaded
 */
function configureLoadStyles(loadStylesFn) {
    _themeState.loadStyles = loadStylesFn;
}
/**
 * Configure run mode of load-themable-styles
 * @param mode load-themable-styles run mode, async or sync
 */
function configureRunMode(mode) {
    _themeState.runState.mode = mode;
}
/**
 * external code can call flush to synchronously force processing of currently buffered styles
 */
function flush() {
    measure(function () {
        var styleArrays = _themeState.runState.buffer.slice();
        _themeState.runState.buffer = [];
        var mergedStyleArray = [].concat.apply([], styleArrays);
        if (mergedStyleArray.length > 0) {
            applyThemableStyles(mergedStyleArray);
        }
    });
}
/**
 * register async loadStyles
 */
function asyncLoadStyles() {
    // Use "self" to distinguish conflicting global typings for setTimeout() from lib.dom.d.ts vs Jest's @types/node
    // https://github.com/jestjs/jest/issues/14418
    return self.setTimeout(function () {
        _themeState.runState.flushTimer = 0;
        flush();
    }, 0);
}
/**
 * Loads a set of style text. If it is registered too early, we will register it when the window.load event
 * is fired.
 * @param {string} styleText Style to register.
 * @param {IStyleRecord} styleRecord Existing style record to re-apply.
 */
function applyThemableStyles(stylesArray, styleRecord) {
    if (_themeState.loadStyles) {
        _themeState.loadStyles(resolveThemableArray(stylesArray).styleString, stylesArray);
    }
    else {
        registerStyles(stylesArray);
    }
}
/**
 * Registers a set theme tokens to find and replace. If styles were already registered, they will be
 * replaced.
 * @param {theme} theme JSON object of theme tokens to values.
 */
function loadTheme(theme) {
    _themeState.theme = theme;
    // reload styles.
    reloadStyles();
}
/**
 * Clear already registered style elements and style records in theme_State object
 * @param option - specify which group of registered styles should be cleared.
 * Default to be both themable and non-themable styles will be cleared
 */
function clearStyles(option) {
    if (option === void 0) { option = ClearStyleOptions.all; }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyNonThemable) {
        clearStylesInternal(_themeState.registeredStyles);
        _themeState.registeredStyles = [];
    }
    if (option === ClearStyleOptions.all || option === ClearStyleOptions.onlyThemable) {
        clearStylesInternal(_themeState.registeredThemableStyles);
        _themeState.registeredThemableStyles = [];
    }
}
function clearStylesInternal(records) {
    records.forEach(function (styleRecord) {
        var styleElement = styleRecord && styleRecord.styleElement;
        if (styleElement && styleElement.parentElement) {
            styleElement.parentElement.removeChild(styleElement);
        }
    });
}
/**
 * Reloads styles.
 */
function reloadStyles() {
    if (_themeState.theme) {
        var themableStyles = [];
        for (var _i = 0, _a = _themeState.registeredThemableStyles; _i < _a.length; _i++) {
            var styleRecord = _a[_i];
            themableStyles.push(styleRecord.themableStyle);
        }
        if (themableStyles.length > 0) {
            clearStyles(ClearStyleOptions.onlyThemable);
            applyThemableStyles([].concat.apply([], themableStyles));
        }
    }
}
/**
 * Find theme tokens and replaces them with provided theme values.
 * @param {string} styles Tokenized styles to fix.
 */
function detokenize(styles) {
    if (styles) {
        styles = resolveThemableArray(splitStyles(styles)).styleString;
    }
    return styles;
}
/**
 * Resolves ThemingInstruction objects in an array and joins the result into a string.
 * @param {ThemableArray} splitStyleArray ThemableArray to resolve and join.
 */
function resolveThemableArray(splitStyleArray) {
    var theme = _themeState.theme;
    var themable = false;
    // Resolve the array of theming instructions to an array of strings.
    // Then join the array to produce the final CSS string.
    var resolvedArray = (splitStyleArray || []).map(function (currentValue) {
        var themeSlot = currentValue.theme;
        if (themeSlot) {
            themable = true;
            // A theming annotation. Resolve it.
            var themedValue = theme ? theme[themeSlot] : undefined;
            var defaultValue = currentValue.defaultValue || 'inherit';
            // Warn to console if we hit an unthemed value even when themes are provided, but only if "DEBUG" is true.
            // Allow the themedValue to be undefined to explicitly request the default value.
            if (theme &&
                !themedValue &&
                console &&
                !(themeSlot in theme) &&
                "boolean" !== 'undefined' &&
                true) {
                // eslint-disable-next-line no-console
                console.warn("Theming value not provided for \"".concat(themeSlot, "\". Falling back to \"").concat(defaultValue, "\"."));
            }
            return themedValue || defaultValue;
        }
        else {
            // A non-themable string. Preserve it.
            return currentValue.rawString;
        }
    });
    return {
        styleString: resolvedArray.join(''),
        themable: themable
    };
}
/**
 * Split tokenized CSS into an array of strings and theme specification objects
 * @param {string} styles Tokenized styles to split.
 */
function splitStyles(styles) {
    var result = [];
    if (styles) {
        var pos = 0; // Current position in styles.
        var tokenMatch = void 0;
        while ((tokenMatch = _themeTokenRegex.exec(styles))) {
            var matchIndex = tokenMatch.index;
            if (matchIndex > pos) {
                result.push({
                    rawString: styles.substring(pos, matchIndex)
                });
            }
            result.push({
                theme: tokenMatch[1],
                defaultValue: tokenMatch[2] // May be undefined
            });
            // index of the first character after the current match
            pos = _themeTokenRegex.lastIndex;
        }
        // Push the rest of the string after the last match.
        result.push({
            rawString: styles.substring(pos)
        });
    }
    return result;
}
/**
 * Registers a set of style text. If it is registered too early, we will register it when the
 * window.load event is fired.
 * @param {ThemableArray} styleArray Array of IThemingInstruction objects to register.
 * @param {IStyleRecord} styleRecord May specify a style Element to update.
 */
function registerStyles(styleArray) {
    if (typeof document === 'undefined') {
        return;
    }
    var head = document.getElementsByTagName('head')[0];
    var styleElement = document.createElement('style');
    var _a = resolveThemableArray(styleArray), styleString = _a.styleString, themable = _a.themable;
    styleElement.setAttribute('data-load-themed-styles', 'true');
    if (_styleNonce) {
        styleElement.setAttribute('nonce', _styleNonce);
    }
    styleElement.appendChild(document.createTextNode(styleString));
    _themeState.perf.count++;
    head.appendChild(styleElement);
    var ev = document.createEvent('HTMLEvents');
    ev.initEvent('styleinsert', true /* bubbleEvent */, false /* cancelable */);
    ev.args = {
        newStyle: styleElement
    };
    document.dispatchEvent(ev);
    var record = {
        styleElement: styleElement,
        themableStyle: styleArray
    };
    if (themable) {
        _themeState.registeredThemableStyles.push(record);
    }
    else {
        _themeState.registeredStyles.push(record);
    }
}


/***/ }),

/***/ 309:
/*!*********************************************************!*\
  !*** ./lib/webparts/helloWorld/assets/welcome-dark.png ***!
  \*********************************************************/
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {

module.exports = __webpack_require__.p + "welcome-dark_bc81978d2f17e05985ee.png";

/***/ }),

/***/ 141:
/*!**********************************************************!*\
  !*** ./lib/webparts/helloWorld/assets/welcome-light.png ***!
  \**********************************************************/
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {

module.exports = __webpack_require__.p + "welcome-light_a2dcb0d64c8d6e80cf49.png";

/***/ }),

/***/ 676:
/*!*********************************************!*\
  !*** external "@microsoft/sp-core-library" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__676__;

/***/ }),

/***/ 529:
/*!**********************************************!*\
  !*** external "@microsoft/sp-lodash-subset" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__529__;

/***/ }),

/***/ 877:
/*!**********************************************!*\
  !*** external "@microsoft/sp-property-pane" ***!
  \**********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__877__;

/***/ }),

/***/ 642:
/*!*********************************************!*\
  !*** external "@microsoft/sp-webpart-base" ***!
  \*********************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__642__;

/***/ }),

/***/ 275:
/*!*******************************************!*\
  !*** external "HelloWorldWebPartStrings" ***!
  \*******************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__275__;

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	(() => {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	(() => {
/******/ 		var _publicPath = __RUSHSTACK_CURRENT_SCRIPT__ ? __RUSHSTACK_CURRENT_SCRIPT__.src : '';
/******/ 		__webpack_require__.p = _publicPath.slice(0, _publicPath.lastIndexOf('/') + 1);
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!******************************************************!*\
  !*** ./lib/webparts/helloWorld/HelloWorldWebPart.js ***!
  \******************************************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-core-library */ 676);
/* harmony import */ var _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! @microsoft/sp-property-pane */ 877);
/* harmony import */ var _microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! @microsoft/sp-webpart-base */ 642);
/* harmony import */ var _microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! @microsoft/sp-lodash-subset */ 529);
/* harmony import */ var _microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var _HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./HelloWorldWebPart.module.scss */ 1);
/* harmony import */ var HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! HelloWorldWebPartStrings */ 275);
/* harmony import */ var HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var _services_TeamsService__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./services/TeamsService */ 266);
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
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







var HelloWorldWebPart = /** @class */ (function (_super) {
    __extends(HelloWorldWebPart, _super);
    function HelloWorldWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        _this._environmentMessage = '';
        _this._clickCount = 0;
        _this._teamsService = null;
        _this._channels = [];
        _this._selectedChannels = new Set();
        _this._isLoading = false;
        _this._error = '';
        _this._channelFiles = {};
        return _this;
    }
    HelloWorldWebPart.prototype.render = function () {
        this.domElement.innerHTML = "\n    <section class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].helloWorld, " ").concat(!!this.context.sdks.microsoftTeams ? _HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].teams : '', "\">\n      <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].headerCard, "\">\n        <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].welcome, "\">\n          <img alt=\"\" src=\"").concat(this._isDarkTheme ? __webpack_require__(/*! ./assets/welcome-dark.png */ 309) : __webpack_require__(/*! ./assets/welcome-light.png */ 141), "\" class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].welcomeImage, "\" />\n          <h2>Well done, ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(this.context.pageContext.user.displayName), "!</h2>\n          <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].environmentInfo, "\">").concat(this._environmentMessage, "</div>\n          <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].propertyInfo, "\">Web part property: <strong>").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(this.properties.description), "</strong></div>\n        </div>\n      </div>\n\n      <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].teamsCard, "\">\n        <h3>\uD83D\uDCCB My Teams Channels</h3>\n        <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].teamsSection, "\">\n          ").concat(this._renderTeamsContent(), "\n        </div>\n      </div>\n\n      <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].interactiveCard, "\">\n        <h3>\uD83C\uDFAF Interactive Test Area</h3>\n        <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].counterSection, "\">\n          <p>Click counter: <span class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].counterDisplay, "\">").concat(this._clickCount, "</span></p>\n          <button class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].primaryButton, "\" data-action=\"increment\">Increment Counter</button>\n          <button class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].secondaryButton, "\" data-action=\"reset\">Reset Counter</button>\n        </div>\n      </div>\n\n      <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].contentCard, "\">\n        <h3>\uD83D\uDCDA Welcome to SharePoint Framework!</h3>\n        <p class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].description, "\">\n        The SharePoint Framework (SPFx) is an extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It's the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.\n        </p>\n        <h4>\uD83D\uDE80 Learn more about SPFx development:</h4>\n          <ul class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].links, "\">\n            <li><a href=\"https://aka.ms/spfx\" target=\"_blank\">\uD83D\uDCD6 SharePoint Framework Overview</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-graph\" target=\"_blank\">\uD83D\uDCCA Use Microsoft Graph in your solution</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-teams\" target=\"_blank\">\uD83D\uDC65 Build for Microsoft Teams using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-viva\" target=\"_blank\">\uD83D\uDCBC Build for Microsoft Viva Connections using SharePoint Framework</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-store\" target=\"_blank\">\uD83C\uDFEA Publish SharePoint Framework applications to the marketplace</a></li>\n            <li><a href=\"https://aka.ms/spfx-yeoman-api\" target=\"_blank\">\uD83D\uDD27 SharePoint Framework API reference</a></li>\n            <li><a href=\"https://aka.ms/m365pnp\" target=\"_blank\">\uFFFD\uD83E\uDD1D Microsoft 365 Developer Community</a></li>\n          </ul>\n      </div>\n    </section>");
        this._bindEvents();
    };
    HelloWorldWebPart.prototype._renderTeamsContent = function () {
        if (this._isLoading) {
            return "\n        <div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].loadingState, "\">\n          <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].spinner, "\"></div>\n          <p>Loading Teams channels...</p>\n        </div>\n      ");
        }
        if (this._error) {
            return "\n        <div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].errorState, "\">\n          <p class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].errorMessage, "\">\u274C ").concat(this._error, "</p>\n          <button class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].primaryButton, "\" data-action=\"retry\">Try Again</button>\n        </div>\n      ");
        }
        if (this._channels.length === 0) {
            return "\n        <div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].emptyState, "\">\n          <button class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].primaryButton, "\" data-action=\"load-channels\">Load My Teams Channels</button>\n        </div>\n      ");
        }
        var groupedChannels = this._groupChannelsByTeam();
        var html = '<div class="${styles.channelsList}">';
        var teamNames = Object.keys(groupedChannels);
        for (var i = 0; i < teamNames.length; i++) {
            var teamName = teamNames[i];
            var channels = groupedChannels[teamName];
            html += "\n        <div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].teamGroup, "\">\n          <h4 class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].teamName, "\">\uD83D\uDC65 ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(teamName), "</h4>\n          <div class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].channelsGroup, "\">\n      ");
            for (var j = 0; j < channels.length; j++) {
                var channel = channels[j];
                var isChecked = this._selectedChannels.has(channel.id);
                var channelKey = "".concat(channel.teamId, "_").concat(channel.id);
                var hasFiles = this._channelFiles[channelKey];
                html += "\n          <div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].channelContainer, "\">\n            <label class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].channelItem, "\">\n              <input type=\"checkbox\"\n                     data-channel-id=\"").concat(channel.id, "\"\n                     ").concat(isChecked ? 'checked' : '', " />\n              <span class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].channelName, "\"># ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(channel.displayName), "</span>\n              ").concat(channel.description ? "<span class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].channelDesc, "\">").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(channel.description), "</span>") : '', "\n              ").concat(!hasFiles ? "<span class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].showFilesBtn, "\" data-team-id=\"").concat(channel.teamId, "\" data-channel-id=\"").concat(channel.id, "\" data-action=\"show-files\">\uD83D\uDCC1</span>") : '', "\n            </label>\n            ").concat(hasFiles ? "<div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].filesList, "\">").concat(this._renderChannelFiles(channelKey), "</div>") : '', "\n          </div>\n        ");
            }
            html += "\n          </div>\n        </div>\n      ";
        }
        html += '</div>';
        return html;
    };
    HelloWorldWebPart.prototype._groupChannelsByTeam = function () {
        return this._channels.reduce(function (groups, channel) {
            var teamName = channel.teamName;
            if (!groups[teamName]) {
                groups[teamName] = [];
            }
            groups[teamName].push(channel);
            return groups;
        }, {});
    };
    HelloWorldWebPart.prototype._bindEvents = function () {
        var _this = this;
        this.domElement.addEventListener('click', function (event) {
            var target = event.target;
            var action = target.getAttribute('data-action');
            if (action === 'increment') {
                _this._clickCount++;
                _this._updateCounter();
            }
            else if (action === 'reset') {
                _this._clickCount = 0;
                _this._updateCounter();
            }
            else if (action === 'load-channels') {
                _this._loadTeamsChannels();
            }
            else if (action === 'retry') {
                _this._error = '';
                _this._loadTeamsChannels();
            }
            else if (action === 'show-files') {
                var teamId = target.getAttribute('data-team-id');
                var channelId = target.getAttribute('data-channel-id');
                if (teamId && channelId) {
                    _this._loadChannelFiles(teamId, channelId);
                }
            }
            else if (action === 'open-folder') {
                var driveId = target.getAttribute('data-drive-id');
                var itemId = target.getAttribute('data-item-id');
                if (driveId && itemId) {
                    _this._loadFolderContents(driveId, itemId, target);
                }
            }
        });
        this.domElement.addEventListener('change', function (event) {
            var target = event.target;
            if (target.type === 'checkbox' && target.dataset.channelId) {
                _this._handleChannelSelection(target.dataset.channelId, target.checked);
            }
        });
    };
    HelloWorldWebPart.prototype._updateCounter = function () {
        var counterDisplay = this.domElement.querySelector(".".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].counterDisplay));
        if (counterDisplay) {
            counterDisplay.textContent = this._clickCount.toString();
        }
    };
    HelloWorldWebPart.prototype._loadTeamsChannels = function () {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, _a, error_1, teamsError;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        this._isLoading = true;
                        this._error = '';
                        this.render();
                        _b.label = 1;
                    case 1:
                        _b.trys.push([1, 5, , 6]);
                        if (!!this._teamsService) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 2:
                        graphClient = _b.sent();
                        this._teamsService = new _services_TeamsService__WEBPACK_IMPORTED_MODULE_6__.TeamsService(graphClient);
                        _b.label = 3;
                    case 3:
                        _a = this;
                        return [4 /*yield*/, this._teamsService.getAllChannelsForUser()];
                    case 4:
                        _a._channels = _b.sent();
                        this._isLoading = false;
                        this.render();
                        return [3 /*break*/, 6];
                    case 5:
                        error_1 = _b.sent();
                        this._isLoading = false;
                        teamsError = error_1;
                        switch (teamsError.type) {
                            case 'NO_PERMISSIONS':
                                this._error = 'Insufficient permissions. Please contact your administrator to approve Microsoft Graph permissions.';
                                break;
                            case 'NO_TEAMS':
                                this._error = 'You are not a member of any teams.';
                                break;
                            case 'NETWORK_ERROR':
                                this._error = 'Network error. Please check your connection and try again.';
                                break;
                            default:
                                this._error = teamsError.message || 'An unexpected error occurred while loading channels.';
                        }
                        this.render();
                        console.error('Error loading Teams channels:', error_1);
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    HelloWorldWebPart.prototype._handleChannelSelection = function (channelId, isSelected) {
        if (isSelected) {
            this._selectedChannels.add(channelId);
        }
        else {
            this._selectedChannels.delete(channelId);
        }
        var selectedChannelsArray = [];
        this._selectedChannels.forEach(function (channelId) { return selectedChannelsArray.push(channelId); });
        console.log('Selected channels:', selectedChannelsArray);
    };
    HelloWorldWebPart.prototype._loadChannelFiles = function (teamId, channelId) {
        return __awaiter(this, void 0, void 0, function () {
            var channelKey, graphClient, files, error_2;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        channelKey = "".concat(teamId, "_").concat(channelId);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 5, , 6]);
                        if (!!this._teamsService) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 2:
                        graphClient = _a.sent();
                        this._teamsService = new _services_TeamsService__WEBPACK_IMPORTED_MODULE_6__.TeamsService(graphClient);
                        _a.label = 3;
                    case 3: return [4 /*yield*/, this._teamsService.loadChannelFilesDirect(teamId, channelId)];
                    case 4:
                        files = _a.sent();
                        this._channelFiles[channelKey] = files;
                        this.render();
                        return [3 /*break*/, 6];
                    case 5:
                        error_2 = _a.sent();
                        console.error('Failed to load channel files:', error_2);
                        return [3 /*break*/, 6];
                    case 6: return [2 /*return*/];
                }
            });
        });
    };
    HelloWorldWebPart.prototype._renderChannelFiles = function (channelKey) {
        var _this = this;
        var files = this._channelFiles[channelKey] || [];
        if (files.length === 0)
            return '<p>No files found</p>';
        var html = '<ul class="' + _HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].filesUL + '">';
        files.forEach(function (file) {
            var icon = file.type === 'folder' ? '📁' : '📄';
            var folderId = "folder_".concat(file.itemId);
            var folderContents = _this._channelFiles[folderId];
            if (file.type === 'folder') {
                html += "<li class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].folderItem, "\">\n          <input type=\"checkbox\" data-file-id=\"").concat(file.id, "\" />\n          <span class=\"").concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].folderName, "\" data-action=\"open-folder\" data-drive-id=\"").concat(file.driveId, "\" data-item-id=\"").concat(file.itemId, "\">").concat(icon, " ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(file.name), "</span>\n        </li>");
                // Show folder contents if loaded
                if (folderContents && folderContents.length > 0) {
                    html += "<li class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].subFolderContent, "\">");
                    folderContents.forEach(function (subFile) {
                        var subIcon = subFile.type === 'folder' ? '📁' : '📄';
                        html += "<div class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].subFileItem, "\">\n              <input type=\"checkbox\" data-file-id=\"").concat(subFile.id, "\" />\n              <span>").concat(subIcon, " ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(subFile.name), "</span>\n            </div>");
                    });
                    html += "</li>";
                }
            }
            else {
                html += "<li class=\"".concat(_HelloWorldWebPart_module_scss__WEBPACK_IMPORTED_MODULE_4__["default"].fileItem, "\">\n          <input type=\"checkbox\" data-file-id=\"").concat(file.id, "\" />\n          <span>").concat(icon, " ").concat((0,_microsoft_sp_lodash_subset__WEBPACK_IMPORTED_MODULE_3__.escape)(file.name), "</span>\n        </li>");
            }
        });
        html += '</ul>';
        return html;
    };
    HelloWorldWebPart.prototype._loadFolderContents = function (driveId, itemId, targetElement) {
        return __awaiter(this, void 0, void 0, function () {
            var graphClient, folderId, folderItems, error_3;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, , 5]);
                        if (!!this._teamsService) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.context.msGraphClientFactory.getClient('3')];
                    case 1:
                        graphClient = _a.sent();
                        this._teamsService = new _services_TeamsService__WEBPACK_IMPORTED_MODULE_6__.TeamsService(graphClient);
                        _a.label = 2;
                    case 2:
                        folderId = "folder_".concat(itemId);
                        // Check if already loaded - toggle off
                        if (this._channelFiles[folderId]) {
                            delete this._channelFiles[folderId];
                            this.render();
                            return [2 /*return*/];
                        }
                        return [4 /*yield*/, this._teamsService.getDriveItemChildren(driveId, itemId)];
                    case 3:
                        folderItems = _a.sent();
                        // Store folder contents and re-render
                        this._channelFiles[folderId] = folderItems.map(function (item) { return ({
                            id: item.id,
                            name: item.name,
                            type: item.folder ? 'folder' : 'file',
                            driveId: driveId,
                            itemId: item.id
                        }); });
                        this.render();
                        return [3 /*break*/, 5];
                    case 4:
                        error_3 = _a.sent();
                        console.error('Failed to load folder contents:', error_3);
                        return [3 /*break*/, 5];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    HelloWorldWebPart.prototype.onInit = function () {
        var _this = this;
        return this._getEnvironmentMessage().then(function (message) {
            _this._environmentMessage = message;
        });
    };
    HelloWorldWebPart.prototype._getEnvironmentMessage = function () {
        var _this = this;
        if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
            return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
                .then(function (context) {
                var environmentMessage = '';
                switch (context.app.host.name) {
                    case 'Office': // running in Office
                        environmentMessage = _this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppLocalEnvironmentOffice : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppOfficeEnvironment;
                        break;
                    case 'Outlook': // running in Outlook
                        environmentMessage = _this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppLocalEnvironmentOutlook : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppOutlookEnvironment;
                        break;
                    case 'Teams': // running in Teams
                    case 'TeamsModern':
                        environmentMessage = _this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppLocalEnvironmentTeams : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppTeamsTabEnvironment;
                        break;
                    default:
                        environmentMessage = HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.UnknownEnvironment;
                }
                return environmentMessage;
            });
        }
        return Promise.resolve(this.context.isServedFromLocalhost ? HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppLocalEnvironmentSharePoint : HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.AppSharePointEnvironment);
    };
    HelloWorldWebPart.prototype.onThemeChanged = function (currentTheme) {
        if (!currentTheme) {
            return;
        }
        this._isDarkTheme = !!currentTheme.isInverted;
        var semanticColors = currentTheme.semanticColors;
        if (semanticColors) {
            this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
            this.domElement.style.setProperty('--link', semanticColors.link || null);
            this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
        }
    };
    Object.defineProperty(HelloWorldWebPart.prototype, "dataVersion", {
        get: function () {
            return _microsoft_sp_core_library__WEBPACK_IMPORTED_MODULE_0__.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    HelloWorldWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.BasicGroupName,
                            groupFields: [
                                (0,_microsoft_sp_property_pane__WEBPACK_IMPORTED_MODULE_1__.PropertyPaneTextField)('description', {
                                    label: HelloWorldWebPartStrings__WEBPACK_IMPORTED_MODULE_5__.DescriptionFieldLabel
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return HelloWorldWebPart;
}(_microsoft_sp_webpart_base__WEBPACK_IMPORTED_MODULE_2__.BaseClientSideWebPart));
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (HelloWorldWebPart);

})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});})();;
//# sourceMappingURL=hello-world-web-part.js.map