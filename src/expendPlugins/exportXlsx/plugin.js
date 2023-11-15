import { writeFileXLSX, utils } from "xlsx";
import locale from "../../locale/locale";
import { modelHTML } from "../../controllers/constant";
import { arrayRemoveItem, replaceHtml } from "../../utils/util";
import tooltip from "../../global/tooltip";
import { getSheetIndex } from "../../methods/get";
import Store from "../../store";
import { getAllSheets } from "../../global/api";

// Initialize the export xlsx api
function exportXlsx(options, config, isDemo) {
  arrayRemoveItem(Store.asyncLoad, "exportXlsx");
}

function downloadXlsx(data, filename) {
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

/**
 *
 * @param {*} url
 * @param {*} success
 * @param {*} fail
 */
function fetchAndDownloadXlsx({ url, order }, success, fail) {
  const luckyJson = luckysheet.toJson();

  luckysheet.getAllChartsBase64(chartMap => {
    luckyJson.chartMap = chartMap;
    luckyJson.devicePixelRatio = window.devicePixelRatio;
    luckyJson.exportXlsx = {
      order,
    };

    fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(luckyJson),
    })
      .then(response => response.blob())
      .then(blob => {
        if (
          blob.type ===
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ) {
          const filename = luckyJson.title + ".xlsx";
          downloadXlsx(blob, filename);
          success && success();
        } else {
          fail && fail();
        }
      })
      .catch(error => {
        console.error("fetch error:", error);
        fail && fail();
      });
  });
}

const formatDataForXLSXPlugin = luckysheet => {
  const result = [];

  for (let r = 0; r < luckysheet.data.length; r++) {
    const row = luckysheet.data[r];
    result[r] = [];

    for (let c = 0; c < row.length; c++) {
      const col = luckysheet.data[r][c];

      if (!col) {
        result[r][c] = "";
      } else {
        result[r][c] = col.v;
      }
    }
  }

  return result;
};

/**
 * @param {string} name
 */
const normalizeFileName = name => {
  return name.trim().replace(/\s/g, "_");
};

function createExportDialog() {
  $("#luckysheet-modal-dialog-mask").hide();
  var xlsxContainer = $("#luckysheet-export-xlsx");

  if (xlsxContainer.length === 0) {
    const _locale = locale();
    const locale_exportXlsx = _locale.exportXlsx;
    const locale_button = _locale.button;

    let content = `<div class="luckysheet-export-xlsx-content" style="padding: 10px 10px 10px 0;">
                <span>${locale_exportXlsx.range}</span>
                <select class="luckysheet-export-xlsx-select-area">
                    <option value="allSheets" selected="selected">${locale_exportXlsx.allSheets}</option>
                    <option value="currentSheet">${locale_exportXlsx.currentSheet}</option>
                </select>
        </div>`;

    $("body").append(
      replaceHtml(modelHTML, {
        id: "luckysheet-export-xlsx",
        addclass: "luckysheet-export-xlsx",
        title: locale_exportXlsx.title,
        content: content,
        botton: `<button class="btn btn-primary luckysheet-model-confirm-btn">${locale_button.confirm}</button><button class="btn btn-default luckysheet-model-close-btn">${locale_button.close}</button>`,
        style: "z-index: 999",
        close: locale_button.close,
      })
    );

    selectedOption = "allSheets";

    // init event
    $("#luckysheet-export-xlsx .luckysheet-model-confirm-btn").on(
      "click",
      () => {
        try {
          // luckysheet.showLoadingProgress();

          // Your two-dimensional array
          let luckysheets = [];

          if (selectedOption === "currentSheet") {
            const allSheets = getAllSheets();
            const currentSheet = allSheets.find(
              sheet => sheet.index == Store.currentSheetIndex
            );
            if (!currentSheet) {
              throw new Error();
            }
            luckysheets.push(currentSheet);
          } else {
            luckysheets = getAllSheets();
          }

          // Create a new workbook
          const workbook = utils.book_new();

          for (const luckysheet of luckysheets) {
            const data = formatDataForXLSXPlugin(luckysheet);

            // Add a worksheet to the workbook
            const worksheet = utils.aoa_to_sheet(data);
            utils.book_append_sheet(workbook, worksheet, luckysheet.name);
          }

          let fileName = "sheet";

          const fileNameForExport = normalizeFileName(Store.fileNameForExport);

          if (fileNameForExport) {
            if (selectedOption === "currentSheet") {
              const currentSheetName = normalizeFileName(luckysheets[0].name);
              fileName = `${fileNameForExport}_${currentSheetName}`;
            } else {
              fileName = fileNameForExport;
            }
          } else if (selectedOption === "currentSheet") {
            const currentSheetName = normalizeFileName(luckysheets[0].name);
            fileName = currentSheetName ? currentSheetName : fileName;
          }

          writeFileXLSX(workbook, `${fileName}.xlsx`);

          $("#luckysheet-export-xlsx").hide();

          // luckysheet.hideLoadingProgress();
        } catch (e) {
          console.log("catch", e);
          tooltip.info(_locale.exportXlsx.unexpectedError, "");
        }
      }
    );

    $("#luckysheet-export-xlsx .luckysheet-export-xlsx-select-area").change(
      function () {
        selectedOption = $(this).val();
      }
    );
  }

  let $t = $("#luckysheet-export-xlsx")
      .find(".luckysheet-modal-dialog-content")
      .css("min-width", 350)
      .end(),
    myh = $t.outerHeight(),
    myw = $t.outerWidth();
  let winw = $(window).width(),
    winh = $(window).height();
  let scrollLeft = $(document).scrollLeft(),
    scrollTop = $(document).scrollTop();
  $("#luckysheet-export-xlsx")
    .css({
      left: (winw + scrollLeft - myw) / 2,
      top: (winh + scrollTop - myh) / 3,
    })
    .show();
}

export { exportXlsx, downloadXlsx, fetchAndDownloadXlsx, createExportDialog };
