const scriptElement = document.createElement('script');
scriptElement.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
document.head.appendChild(scriptElement);

const scriptElementUI = document.createElement('script');
scriptElementUI.src = 'https://unpkg.com/kintone-ui-component/umd/kuc.min.js';
document.head.appendChild(scriptElementUI);

(function () {
    'use strict';

    function processAttachments(recordId) {
        const body = {
            'app': kintone.app.getId(),
            'id': recordId
        };

        return kintone.api(kintone.api.url('/k/v1/record.json', true), 'GET', body)
            .then(resp => {
                const attachments = resp.record['Attachment'].value;
                const POReceivedDate = resp.record['Received_Date'].value;

                return Promise.all(attachments.map(attachment => {
                    const fileKey = attachment.fileKey;
                    const url = `https://stevendemo.kintone.com/k/v1/file.json?fileKey=${fileKey}`;

                    return new Promise((resolve, reject) => {
                        const xhr = new XMLHttpRequest();
                        xhr.open('GET', url);
                        xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
                        xhr.responseType = 'blob';

                        xhr.onload = function () {
                            if (xhr.status === 200) {
                                const blob = xhr.response;

                                const reader = new FileReader();
                                reader.onload = function (event) {
                                    const data = event.target.result;
                                    const workbook = XLSX.read(data, { type: 'binary' });
                                    const sheetName = workbook.SheetNames[0];
                                    const worksheet = workbook.Sheets[sheetName];

                                    const CellValuePONumber = worksheet['J3'].v;
                                    const cellValueStartDate = worksheet['A9'].v;
                                    const cellValueCompanyDetail = worksheet['H7'].v;
                                    const cellValueShippedVia = worksheet['E9'].v;
                                    const cellValueFOB = worksheet['H9'].v;
                                    const cellValueSubTotal = worksheet['J36'].v;
                                    const cellValueSalesTax = worksheet['J37'].v;
                                    const cellValueTotalAmount = worksheet['J38'].v;

                                    const tableData = [];
                                    for (let i = 11; i <= 16; i++) {
                                        const tableCellValueA = worksheet['A' + i].v;
                                        const tableCellValueB = worksheet['B' + i].v;
                                        const tableCellValueH = worksheet['H' + i].v;
                                        const tableCellValueJ = worksheet['J' + i].v;
                                        tableData.push({
                                            value: {
                                                "Unit": { "value": tableCellValueA },
                                                "Description": { "value": tableCellValueB },
                                                "Unit_Price": { "value": tableCellValueH },
                                                "Amount": { "value": tableCellValueJ }
                                            }
                                        });
                                    }

                                    const dateArray = cellValueStartDate.split('/');
                                    const formattedDateCellValue = `${dateArray[2]}-${dateArray[0]}-${dateArray[1]}`;

                                    const postData = {
                                        'app': 1589,
                                        'record': {
                                            'PO_Number': { 'value': CellValuePONumber },
                                            'Start_Date': { type: 'DATE', value: formattedDateCellValue },
                                            'Table': { 'value': tableData },
                                            'Company_Detail': { 'value': cellValueCompanyDetail },
                                            'Shipped_via': { 'value': cellValueShippedVia },
                                            'FOB': { 'value': cellValueFOB },
                                            'Subtotal': { 'value': cellValueSubTotal },
                                            'Sales_Tax': { 'value': cellValueSalesTax },
                                            'Total_Amount': { 'value': cellValueTotalAmount },
                                            'PO_Received': { 'value': POReceivedDate },
                                        }
                                    };

                                    kintone.api(kintone.api.url('/k/v1/record', true), 'POST', postData)
                                        .then(postResp => {
                                            alert("The data succesfully extracted and sent to another app")
                                            const updateRecordData = {
                                                'app': kintone.app.getId(),
                                                'id': recordId,
                                                'record': {
                                                    'Status_0': { 'value': 'File Extracted' }
                                                }
                                            };

                                            kintone.api(kintone.api.url('/k/v1/record.json', true), 'PUT', updateRecordData)
                                                .then(updateResp => {
                                                    console.log('Status updated successfully', updateResp);
                                                }).catch(updateErr => {
                                                    console.error('Failed to update status', updateErr);
                                                });
                                        }).catch(err => {
                                            console.error('Failed to post data to another app', err);
                                            reject(err);
                                        });
                                };
                                reader.readAsBinaryString(blob);

                            } else {
                                console.error('Failed to download file. Status:', xhr.status);
                                reject(new Error('Failed to download file'));
                            }
                        };

                        xhr.onerror = function () {
                            console.error('XHR error occurred');
                            reject(new Error('XHR error'));
                        };

                        xhr.send();
                    });
                }));

            }).catch(err => {
                console.error("Error during data obtaining!", err);
                alert(`Error during data obtaining! - error: ${err.message}`);
            });
    }

    kintone.events.on('app.record.detail.show', function (event) {
        const Kuc = Kucs['1.17.1'];
        const button = new Kuc.Button({
            text: 'Extract Data',
            type: 'normal'
        });

        button.addEventListener('click', function () {
            processAttachments(event.recordId);

        });
        kintone.app.record.getHeaderMenuSpaceElement().appendChild(button);

        return event;
    });

})();
