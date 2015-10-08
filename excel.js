angular.module('elandNrslClientApp')
.factory('excelService', function ($window, $q, $http, resourceServerUrl) {
    var msie = $window.navigator.userAgent.indexOf("MSIE ");
    var sa = null;
    var txtArea1 = null;
    var tab_text = "";
    var tabArray = [];
    var maxWidth = 0;

    var getContents = function (tab, rownum) {

        var base64_array = [];
        var base64_array_int = [];
        var targetImgs = tab.rows[rownum].getElementsByTagName('img'); // Image 엑셀에서도 크기 적용되도록 기본 속성에 height/width 추가

        return $q(function (resolve, reject) { //tab, rownum

            if (targetImgs.length != 0) { // 내용에 이미지가 존재한다면

                for (j = 0; j < targetImgs.length; j++) {
                    
                    // 가장 긴 이미지 Width 값 및 이미지 Height 설정
                    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) { //IE9
                        targetImgs[j].setAttribute('height', targetImgs[j].naturalHeight);
                        if (maxWidth < targetImgs[j].naturalWidth) {
                            maxWidth = targetImgs[j].naturalWidth;
                        };
                    } else {
                        targetImgs[j].setAttribute('height', targetImgs[j].height); //크롬
                        if (maxWidth < targetImgs[j].width) { 
                            maxWidth = targetImgs[j].width;
                        };
                    }

                    targetImgs[j].outerHTML += "<br />"; // 텍스트와 이미지가 한 라인에 있는 경우 엑셀에서 이미지와 텍스트가 겹치므로, 임의로 한 줄 넘겨준다.

                    if (targetImgs[j].src.match(/base64/gi) != null) {
                        base64_array_int.push(j);
                        base64_array.push(targetImgs[j].src);
                    };

                }

                if (base64_array.length != 0) {
                    $http({
                        method: 'post',
                        url: resourceServerUrl.requestUrl + '/api/sample/excel/base64',
                        headers: {
                            'Content-type': 'application/json'
                        },
                        data: { 'base64Img': base64_array }
                    })
                      .success(function (data, status, headers, config) {
                          for (var i = 0; i < base64_array_int.length; i++) {
                              targetImgs[base64_array_int[i]].src = data[i];
                          }
                          tab_text = tab_text + "<tr>" + tab.rows[rownum].innerHTML + "</tr>";

                          resolve(tab_text);
                      })
                    .error(function () {
                        reject("엑셀 다운로드 중 오류가 발생하였습니다.");
                    });
                } else {
                    tab_text = tab_text + "<tr>" + tab.rows[rownum].innerHTML + "</tr>";
                    resolve(tab_text);
                }


            } else {
                tab_text = tab_text + "<tr>" + tab.rows[rownum].innerHTML + "</tr>";
                resolve(tab_text);
            }

        });
       
    };

    return {
        fnExcelReport: function (tab) {
            // 초기화
            maxWidth = 0;
            sa = null;
            tab_text = "";
            var rowLength = tab.rows.length - 1;
            var deferred = $q.defer(); // promise

            for (i = 0 ; i < rowLength; i++) { // 내용 이전 row 처리
                tab_text = tab_text + "<tr>" + tab.rows[i].innerHTML + "</tr>";
            }


            getContents(tab, rowLength).then( // promise
                    function (result) {

                        if (maxWidth != 0) {
                            result = "<colgroup><col style=\'width:100px;\'></col><col width=\'" + maxWidth + "\'></col></colgroup>" + result;
                            result = "<table border='1' width=\'" + (maxWidth + 500) + "\'>" + result + "</table>";
                        } else {
                            result = "<colgroup><col style=\'width:100px;\'></col><col width=\'" + maxWidth + "\'></col></colgroup>" + result;
                            result = "<table border='1'>" + result + "</table>";
                        }

                        
                        result = result.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params


                        // download excel
                        if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // if Internet Explorer
                        {
                            txtArea1 = $window.open();
                            self.focus();
                            txtArea1.document.open("txt/html", "replace");
                            txtArea1.document.write(result);
                            txtArea1.document.close();
                            txtArea1.focus();
                            sa = txtArea1.document.execCommand("SaveAs", true, "excel.xls");
                        }
                        else { // other browsers
                            result = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" + result;
                            $window.btoa(unescape(encodeURIComponent(result))) // encode
                            sa = $window.open('data:application/vnd.ms-excel,' + encodeURIComponent(result));
                        }

                        if (txtArea1) {
                            txtArea1.close();
                        }

                        return (sa);

                    },
                    function (error) { console.log(error); }
                );

            
        }


        //,fnExcelReportServer: function (tab) {
        //    for (i = 0 ; i < tab.rows.length ; i++) {
        //        var targetImgs = tab.rows[i].getElementsByTagName('img'); // Image 엑셀에서도 크기 적용되도록 기본 속성에 height/width 추가
        //        if (targetImgs.length != 0) {
        //            for (j = 0; j < targetImgs.length; j++) {
        //                targetImgs[j].setAttribute('height', targetImgs[j].height);
        //                targetImgs[j].outerHTML += "<br />"; // 텍스트와 이미지가 한 라인에 있는 경우 엑셀에서 이미지와 텍스트가 겹치므로, 임의로 한 줄 넘겨준다.
        //                if (maxWidth < targetImgs[j].width) { // 가장 긴 이미지 Width 값 얻어 셀 크기 결정
        //                    maxWidth = targetImgs[j].width;
        //                }
        //            }
        //        }

        //        tabArray.push(tab.rows[i].innerHTML);
        //    }

        //    $http({
        //        method: 'post',
        //        url: 'http://localhost:3050/api/sample/excel',
        //        headers: {
        //            'Content-type': 'application/json'
        //        },
        //        responseType: 'arraybuffer',
        //        data: { 'base64Img': tabArray }
        //    }).success(function (data, status, headers, config) {
        //        var blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        //        var objectUrl = URL.createObjectURL(blob);
        //        var anchor = angular.element('<a/>');
        //        anchor.attr({
        //            href: objectUrl,
        //            target: '_blank',
        //            download: 'excelList.xlsx'
        //        })[0].click();
        //        anchor.remove();
        //    })
        //      .error(function () {
        //          alert("파일 업로드 중 오류가 발생하였습니다.");
        //      });


        //},
        //fnExcelReport: function (tab) {
        //    for (i = 0 ; i < tab.rows.length ; i++) {
        //        var targetImgs = tab.rows[i].getElementsByTagName('img'); // Image 엑셀에서도 크기 적용되도록 기본 속성에 height/width 추가
        //        if (targetImgs.length != 0) {
        //            for (j = 0; j < targetImgs.length; j++) {
        //                targetImgs[j].setAttribute('height', targetImgs[j].height);
        //                targetImgs[j].outerHTML += "<br />"; // 텍스트와 이미지가 한 라인에 있는 경우 엑셀에서 이미지와 텍스트가 겹치므로, 임의로 한 줄 넘겨준다.
        //                if (maxWidth < targetImgs[j].width) { // 가장 긴 이미지 Width 값 얻어 셀 크기 결정
        //                    maxWidth = targetImgs[j].width;
        //                }
        //            }
        //        }

        //        tab_text = tab_text + "<tr>" + tab.rows[i].innerHTML + "</tr>";
        //        tabArray.push(tab.rows[i].innerHTML);
        //    }
        //    if (maxWidth != 0) {
        //        tab_text = "<colgroup><col style=\'width:100px;\'></col><col width=\'" + maxWidth + "\'></col></colgroup>" + tab_text;
        //    }
        //    tab_text = "<table border='1'>" + tab_text + "</table>";
        //    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params
        //    //test_text = tab_text.match(/<img[^>]*>|<\/img>/gi);
        //    //var base64Img = (test_text[0].split("=")[1] + "\=").split(",")[1].toString();

        //    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./))      // if Internet Explorer
        //    {
        //        txtArea1 = $window.open();
        //        self.focus();
        //        txtArea1.document.open("txt/html", "replace");
        //        txtArea1.document.write(tab_text);
        //        txtArea1.document.close();
        //        txtArea1.focus();
        //        sa = txtArea1.document.execCommand("SaveAs", true, "excel.xls");
        //    }
        //    else { // other browsers
        //        tab_text = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" + tab_text;
        //        $window.btoa(unescape(encodeURIComponent(tab_text))) // encode
        //        sa = $window.open('data:application/vnd.ms-excel,' + encodeURIComponent(tab_text));
        //    }

        //    if (txtArea1) {
        //        txtArea1.close();
        //    }

        //    return (sa);
        //}
    };
})

