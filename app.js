const { createApp, ref, reactive, computed } = Vue;

const app = createApp({
    setup() {
        const currentStep = ref(0);
        const isDragOver = ref(false);
        const uploadError = ref('');
        const fileName = ref('');
        const previewLoading = ref(false);
        const exporting = ref(false);
        const exportProgress = ref(0);
        const exportSummary = ref('');
        const showOnlyWarnings = ref(false);
        const previewImageUrl = ref(null);
        const fileInput = ref(null);

        const workbook = ref(null);
        const sheetList = ref([]);
        const parsedImages = ref({});
        const parsedCells = ref({});

        const config = reactive({
            sheet: '',
            imageCol: 1,
            imageRow: 1,
            nameCol: 2,
            strategy: 'B'
        });

        const previewData = ref([]);
        const logData = ref([]);
        const usedNamesMap = ref({});

        const filteredPreviewData = computed(() => {
            if (!showOnlyWarnings.value) return previewData.value;
            return previewData.value.filter(r => r.status === '警告' || r.status === '错误');
        });

        const selectedCount = computed(() => {
            return previewData.value.filter(r => r.selected).length;
        });

        const allSelected = computed(() => {
            return previewData.value.length > 0 && previewData.value.every(r => r.selected);
        });

        const triggerFileInput = () => {
            fileInput.value?.click();
        };

        const handleFileSelect = (e) => {
            const file = e.target.files[0];
            if (file) processFile(file);
        };

        const handleDrop = (e) => {
            isDragOver.value = false;
            const file = e.dataTransfer.files[0];
            if (file) processFile(file);
        };

        const processFile = async (file) => {
            uploadError.value = '';
            fileName.value = '';

            if (!file.name.endsWith('.xlsx')) {
                uploadError.value = '仅支持 .xlsx 格式文件，不支持 .xls';
                return;
            }

            if (file.size > 100 * 1024 * 1024) {
                uploadError.value = '文件大小不能超过 100MB';
                return;
            }

            try {
                const arrayBuffer = await file.arrayBuffer();
                workbook.value = XLSX.read(arrayBuffer, { type: 'array' });

                sheetList.value = workbook.value.SheetNames;
                if (sheetList.value.length === 0) {
                    uploadError.value = 'Excel文件中没有工作表';
                    return;
                }

                await parseWorkbook(arrayBuffer);

                config.sheet = sheetList.value[0];
                fileName.value = file.name;
                currentStep.value = 1;
            } catch (err) {
                uploadError.value = '文件解析失败: ' + (err.message || '文件可能已损坏或不是有效的Excel文件');
                console.error(err);
            }
        };

        const parseWorkbook = async (arrayBuffer) => {
            parsedImages.value = {};
            parsedCells.value = {};

            workbook.value.SheetNames.forEach(sheetName => {
                const ws = workbook.value.Sheets[sheetName];
                parsedCells.value[sheetName] = {};
                parsedImages.value[sheetName] = [];

                for (const cellId in ws) {
                    if (cellId.startsWith('!')) continue;
                    if (typeof ws[cellId] === 'object' && 'v' in ws[cellId]) {
                        parsedCells.value[sheetName][cellId] = ws[cellId].v;
                    }
                }
            });

            const zip = await JSZip.loadAsync(arrayBuffer);

            const relsMap = {};
            workbook.value.SheetNames.forEach((sheetName, idx) => {
                const relFile = `xl/worksheets/_rels/sheet${idx + 1}.xml.rels`;
                if (zip.file(relFile)) {
                    zip.file(relFile).async('text').then(xmlContent => {
                        const parser = new DOMParser();
                        const doc = parser.parseFromString(xmlContent, 'text/xml');
                        const rels = doc.getElementsByTagName('Relationship');
                        for (let i = 0; i < rels.length; i++) {
                            const rId = rels[i].getAttribute('Id');
                            const target = rels[i].getAttribute('Target');
                            if (target && target.includes('drawing')) {
                                relsMap[sheetName] = relsMap[sheetName] || {};
                                relsMap[sheetName][rId] = target.replace('../', 'xl/');
                            }
                        }
                    });
                }
            });

            const drawingRelsMap = {};
            const drawingFiles = zip.file(/^xl\/drawings\/_rels\/drawing\d+\.xml\.rels$/);
            for (const relFile of drawingFiles) {
                const xmlContent = await relFile.async('text');
                const parser = new DOMParser();
                const doc = parser.parseFromString(xmlContent, 'text/xml');
                const rels = doc.getElementsByTagName('Relationship');
                const drawingName = relFile.name.replace('_rels/', '').replace('.rels', '');
                drawingRelsMap[drawingName] = {};
                for (let i = 0; i < rels.length; i++) {
                    const rId = rels[i].getAttribute('Id');
                    const target = rels[i].getAttribute('Target');
                    if (target) {
                        drawingRelsMap[drawingName][rId] = target.replace('../', 'xl/');
                    }
                }
            }

            const mediaMap = {};
            const mediaFiles = zip.file(/^xl\/media\//);
            for (const mediaFile of mediaFiles) {
                const fName = mediaFile.name.split('/').pop();
                const ext = fName.split('.').pop().toLowerCase();
                if (['png', 'jpg', 'jpeg', 'gif', 'bmp'].includes(ext)) {
                    const ab = await mediaFile.async('arraybuffer');
                    const base64 = arrayBufferToBase64(ab);
                    const mime = getMimeType(ext);
                    mediaMap[mediaFile.name] = {
                        data: ab,
                        thumbnail: `data:${mime};base64,${base64}`,
                        extension: ext
                    };
                }
            }

            for (const sheetName of workbook.value.SheetNames) {
                const sheetIdx = workbook.value.SheetNames.indexOf(sheetName);
                const drawingFile = `xl/drawings/drawing${sheetIdx + 1}.xml`;

                if (!zip.file(drawingFile)) continue;

                const xmlContent = await zip.file(drawingFile).async('text');
                const parser = new DOMParser();
                const doc = parser.parseFromString(xmlContent, 'text/xml');

                const twoCellAnchors = doc.getElementsByTagName('xdr:twoCellAnchor');
                for (let i = 0; i < twoCellAnchors.length; i++) {
                    const anchor = twoCellAnchors[i];
                    const from = anchor.getElementsByTagName('xdr:from')[0];
                    const to = anchor.getElementsByTagName('xdr:to')[0];

                    if (!from || !to) continue;

                    const startCol = parseInt(getXmlElementText(from, 'xdr:col')) + 1;
                    const startRow = parseInt(getXmlElementText(from, 'xdr:row')) + 1;
                    const endCol = parseInt(getXmlElementText(to, 'xdr:col')) + 1;
                    const endRow = parseInt(getXmlElementText(to, 'xdr:row')) + 1;

                    const blip = anchor.getElementsByTagName('a:blip')[0];
                    const embedId = blip?.getAttribute('r:embed');
                    if (!embedId) continue;

                    let mediaPath = null;
                    const drawingRels = drawingRelsMap[drawingFile];
                    if (drawingRels && drawingRels[embedId]) {
                        mediaPath = drawingRels[embedId];
                    }

                    let imageInfo = null;
                    if (mediaPath && mediaMap[mediaPath]) {
                        imageInfo = mediaMap[mediaPath];
                    } else {
                        for (const path in mediaMap) {
                            if (path.includes(embedId.replace('rId', ''))) {
                                imageInfo = mediaMap[path];
                                break;
                            }
                        }
                    }

                    if (!imageInfo) {
                        for (const path in mediaMap) {
                            imageInfo = mediaMap[path];
                            break;
                        }
                    }

                    if (imageInfo) {
                        parsedImages.value[sheetName].push({
                            id: embedId,
                            startCol,
                            startRow,
                            endCol,
                            endRow,
                            thumbnail: imageInfo.thumbnail,
                            data: imageInfo.data,
                            extension: imageInfo.extension
                        });
                    }
                }

                const oneCellAnchors = doc.getElementsByTagName('xdr:oneCellAnchor');
                for (let i = 0; i < oneCellAnchors.length; i++) {
                    const anchor = oneCellAnchors[i];
                    const from = anchor.getElementsByTagName('xdr:from')[0];
                    if (!from) continue;

                    const startCol = parseInt(getXmlElementText(from, 'xdr:col')) + 1;
                    const startRow = parseInt(getXmlElementText(from, 'xdr:row')) + 1;

                    const blip = anchor.getElementsByTagName('a:blip')[0];
                    const embedId = blip?.getAttribute('r:embed');
                    if (!embedId) continue;

                    let mediaPath = null;
                    const drawingRels = drawingRelsMap[drawingFile];
                    if (drawingRels && drawingRels[embedId]) {
                        mediaPath = drawingRels[embedId];
                    }

                    let imageInfo = null;
                    if (mediaPath && mediaMap[mediaPath]) {
                        imageInfo = mediaMap[mediaPath];
                    }

                    if (imageInfo) {
                        parsedImages.value[sheetName].push({
                            id: embedId,
                            startCol,
                            startRow,
                            endCol: startCol + 1,
                            endRow: startRow + 1,
                            thumbnail: imageInfo.thumbnail,
                            data: imageInfo.data,
                            extension: imageInfo.extension
                        });
                    }
                }
            }
        };

        const getXmlElementText = (parent, tagName) => {
            const el = parent.getElementsByTagName(tagName)[0];
            return el ? el.textContent : '0';
        };

        const arrayBufferToBase64 = (buffer) => {
            let binary = '';
            const bytes = new Uint8Array(buffer);
            const chunkSize = 8192;
            for (let i = 0; i < bytes.length; i += chunkSize) {
                const chunk = bytes.subarray(i, i + chunkSize);
                binary += String.fromCharCode.apply(null, chunk);
            }
            return btoa(binary);
        };

        const getMimeType = (ext) => {
            const map = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };
            return map[ext] || 'image/png';
        };

        const colNumToLetter = (num) => {
            let letter = '';
            while (num > 0) {
                num--;
                letter = String.fromCharCode(65 + (num % 26)) + letter;
                num = Math.floor(num / 26);
            }
            return letter;
        };

        const sanitizeFileName = (name) => {
            return String(name).replace(/[\\/:*?"<>|]/g, '_').substring(0, 200);
        };

        const ensureUniqueName = (baseName, ext, usedMap) => {
            const key = `${baseName}.${ext}`;
            if (!usedMap[key]) {
                usedMap[key] = 1;
                return key;
            }
            const newName = `${baseName}_${usedMap[key]}.${ext}`;
            usedMap[key]++;
            return newName;
        };

        const onSheetChange = () => {
            previewData.value = [];
        };

        const generatePreview = () => {
            if (!config.sheet) return;
            previewLoading.value = true;

            setTimeout(() => {
                try {
                    const sheetName = config.sheet;
                    const images = parsedImages.value[sheetName] || [];
                    const cells = parsedCells.value[sheetName] || {};

                    if (images.length === 0) {
                        alert('该工作表中未检测到任何图片，请检查文件');
                        previewLoading.value = false;
                        return;
                    }

                    previewData.value = [];
                    logData.value = [];
                    const usedMap = {};

                    images.forEach((image, imageIndex) => {
                        const imageId = `IMG_${imageIndex + 1}`;

                        if (config.strategy === 'A') {
                            const nameCellId = `${colNumToLetter(config.nameCol)}${image.startRow}`;
                            const nameText = cells[nameCellId] != null ? String(cells[nameCellId]) : '';
                            const isNameEmpty = nameText.trim() === '';

                            let finalBaseName = isNameEmpty ? `unnamed_行${image.startRow}` : sanitizeFileName(nameText);
                            let finalName = ensureUniqueName(finalBaseName, image.extension, usedMap);

                            const status = isNameEmpty ? '警告' : '正常';
                            const reason = isNameEmpty ? '名字单元格为空，使用默认名' : '';

                            previewData.value.push({
                                selected: true,
                                thumbnail: image.thumbnail,
                                range: `行${image.startRow}-${image.endRow}`,
                                nameSource: `行${image.startRow},${colNumToLetter(config.nameCol)}列`,
                                nameText: nameText,
                                finalName: finalName,
                                status: status,
                                image: image,
                                imageId: imageId
                            });

                            logData.value.push({
                                imageId: imageId,
                                range: `行${image.startRow}-${image.endRow}, 列${colNumToLetter(image.startCol)}-${colNumToLetter(image.endCol)}`,
                                nameSource: `行${image.startRow},${colNumToLetter(config.nameCol)}列`,
                                finalName: finalName,
                                status: status,
                                reason: reason || '成功'
                            });

                        } else if (config.strategy === 'B') {
                            for (let row = image.startRow; row <= image.endRow; row++) {
                                const nameCellId = `${colNumToLetter(config.nameCol)}${row}`;
                                const nameText = cells[nameCellId] != null ? String(cells[nameCellId]) : '';
                                const isNameEmpty = nameText.trim() === '';

                                let finalBaseName = isNameEmpty ? `unnamed_行${row}` : sanitizeFileName(nameText);
                                let finalName = ensureUniqueName(finalBaseName, image.extension, usedMap);

                                const status = isNameEmpty ? '警告' : '正常';
                                const reason = isNameEmpty ? '名字单元格为空，使用默认名' : '';

                                previewData.value.push({
                                    selected: true,
                                    thumbnail: image.thumbnail,
                                    range: `行${image.startRow}-${image.endRow}`,
                                    nameSource: `行${row},${colNumToLetter(config.nameCol)}列`,
                                    nameText: nameText,
                                    finalName: finalName,
                                    status: status,
                                    image: image,
                                    imageId: `${imageId}_行${row}`
                                });

                                logData.value.push({
                                    imageId: `${imageId}_行${row}`,
                                    range: `行${image.startRow}-${image.endRow}, 列${colNumToLetter(image.startCol)}-${colNumToLetter(image.endCol)}`,
                                    nameSource: `行${row},${colNumToLetter(config.nameCol)}列`,
                                    finalName: finalName,
                                    status: status,
                                    reason: reason || '成功'
                                });
                            }

                        } else if (config.strategy === 'C') {
                            for (let col = image.startCol; col <= image.endCol; col++) {
                                const nameCellId = `${colNumToLetter(col)}${image.startRow}`;
                                const nameText = cells[nameCellId] != null ? String(cells[nameCellId]) : '';
                                const isNameEmpty = nameText.trim() === '';

                                let finalBaseName = isNameEmpty ? `unnamed_列${colNumToLetter(col)}` : sanitizeFileName(nameText);
                                let finalName = ensureUniqueName(finalBaseName, image.extension, usedMap);

                                const status = isNameEmpty ? '警告' : '正常';
                                const reason = isNameEmpty ? '名字单元格为空，使用默认名' : '';

                                previewData.value.push({
                                    selected: true,
                                    thumbnail: image.thumbnail,
                                    range: `列${colNumToLetter(image.startCol)}-${colNumToLetter(image.endCol)}`,
                                    nameSource: `${colNumToLetter(col)}列,行${image.startRow}`,
                                    nameText: nameText,
                                    finalName: finalName,
                                    status: status,
                                    image: image,
                                    imageId: `${imageId}_列${colNumToLetter(col)}`
                                });

                                logData.value.push({
                                    imageId: `${imageId}_列${colNumToLetter(col)}`,
                                    range: `行${image.startRow}-${image.endRow}, 列${colNumToLetter(image.startCol)}-${colNumToLetter(image.endCol)}`,
                                    nameSource: `${colNumToLetter(col)}列,行${image.startRow}`,
                                    finalName: finalName,
                                    status: status,
                                    reason: reason || '成功'
                                });
                            }
                        }
                    });

                    if (previewData.value.length === 0) {
                        alert('根据配置未生成任何图片副本，请检查配置');
                    } else {
                        currentStep.value = 2;
                    }
                } catch (err) {
                    alert('预览生成失败: ' + err.message);
                    console.error(err);
                } finally {
                    previewLoading.value = false;
                }
            }, 100);
        };

        const selectAll = () => {
            previewData.value.forEach(r => r.selected = true);
        };

        const deselectAll = () => {
            previewData.value.forEach(r => r.selected = false);
        };

        const toggleAll = (e) => {
            const checked = e.target.checked;
            previewData.value.forEach(r => r.selected = checked);
        };

        const previewImage = (url) => {
            previewImageUrl.value = url;
        };

        const exportSingle = (row) => {
            if (!row.image || !row.image.data) {
                alert('图片数据异常，无法导出');
                return;
            }
            try {
                const mime = getMimeType(row.image.extension);
                const blob = new Blob([row.image.data], { type: mime });
                saveAs(blob, row.finalName);
            } catch (err) {
                alert('导出失败: ' + err.message);
            }
        };

        const exportSelected = async () => {
            const selectedItems = previewData.value.filter(r => r.selected);
            if (selectedItems.length === 0) return;

            exporting.value = true;
            exportProgress.value = 0;
            exportSummary.value = '';

            try {
                const zip = new JSZip();
                const usedFileNames = {};
                let successCount = 0;
                let failCount = 0;

                for (let i = 0; i < selectedItems.length; i++) {
                    const item = selectedItems[i];

                    if (!item.image || !item.image.data) {
                        failCount++;
                        logData.value.push({
                            imageId: item.imageId,
                            range: item.range,
                            nameSource: item.nameSource,
                            finalName: item.finalName,
                            status: '错误',
                            reason: '图片数据异常，无法导出'
                        });
                        continue;
                    }

                    let fileName = item.finalName;
                    fileName = fileName.replace(/\.\./g, '_').replace(/\//g, '_');

                    if (usedFileNames[fileName]) {
                        const dotIdx = fileName.lastIndexOf('.');
                        const base = fileName.substring(0, dotIdx);
                        const ext = fileName.substring(dotIdx);
                        fileName = `${base}_${usedFileNames[fileName]}${ext}`;
                        usedFileNames[fileName] = (usedFileNames[fileName] || 1) + 1;
                    } else {
                        usedFileNames[fileName] = 1;
                    }

                    zip.file(fileName, item.image.data);
                    successCount++;

                    exportProgress.value = Math.round(((i + 1) / selectedItems.length) * 90);
                }

                exportProgress.value = 95;
                const content = await zip.generateAsync({
                    type: 'blob',
                    compression: 'DEFLATE',
                    compressionOptions: { level: 6 }
                });

                saveAs(content, 'excel_images.zip');
                exportProgress.value = 100;

                const warningCount = selectedItems.filter(r => r.status === '警告').length;
                exportSummary.value = `导出完成！成功: ${successCount} | 失败: ${failCount} | 警告(使用默认名): ${warningCount}`;
                currentStep.value = 3;
            } catch (err) {
                alert('导出失败: ' + err.message);
                console.error(err);
            } finally {
                exporting.value = false;
            }
        };

        const downloadLog = () => {
            if (logData.value.length === 0) return;

            const headers = ['图片ID', '覆盖范围', '名字来源', '最终文件名', '状态', '详细原因'];
            const csvRows = [
                headers.join(','),
                ...logData.value.map(row =>
                    [row.imageId, row.range, row.nameSource, row.finalName, row.status, row.reason]
                        .map(f => `"${String(f || '').replace(/"/g, '""')}"`)
                        .join(',')
                )
            ];

            const blob = new Blob(['\ufeff' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
            saveAs(blob, 'export_log.csv');
        };

        const goStep = (step) => {
            if (step === 0 && currentStep.value > 0) {
                if (!confirm('返回将丢失当前数据，确定吗？')) return;
            }
            if (step === 1 && !workbook.value) return;
            if (step === 2 && previewData.value.length === 0) return;
            if (step === 3 && selectedCount.value === 0) return;
            currentStep.value = step;
        };

        const resetAll = () => {
            if (!confirm('确定要重置所有数据吗？')) return;
            currentStep.value = 0;
            workbook.value = null;
            sheetList.value = [];
            parsedImages.value = {};
            parsedCells.value = {};
            previewData.value = [];
            logData.value = [];
            fileName.value = '';
            uploadError.value = '';
            exportSummary.value = '';
            exportProgress.value = 0;
            config.sheet = '';
            config.imageCol = 1;
            config.imageRow = 1;
            config.nameCol = 2;
            config.strategy = 'B';
        };

        return {
            currentStep,
            isDragOver,
            uploadError,
            fileName,
            previewLoading,
            exporting,
            exportProgress,
            exportSummary,
            showOnlyWarnings,
            previewImageUrl,
            fileInput,
            sheetList,
            config,
            previewData,
            logData,
            filteredPreviewData,
            selectedCount,
            allSelected,
            triggerFileInput,
            handleFileSelect,
            handleDrop,
            onSheetChange,
            generatePreview,
            selectAll,
            deselectAll,
            toggleAll,
            previewImage,
            exportSingle,
            exportSelected,
            downloadLog,
            goStep,
            resetAll
        };
    }
});

app.mount('#app');
