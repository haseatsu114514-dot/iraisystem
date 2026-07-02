function generateDocuments_(rowData, caseId, caseFolder, validation, force) {
  const templates = getEnabledTemplates_();
  if (!templates.length) {
    throw new Error('有効なテンプレートがありません。初期設定を実行してください。');
  }

  const data = buildTemplateData_(rowData.submission, caseId, validation);
  const results = [];
  templates.forEach(function(template) {
    const suffix = force
      ? '_再生成_' + Utilities.formatDate(new Date(), ASTRA_CONFIG.TIME_ZONE, 'yyyyMMdd_HHmmss')
      : '';
    const fileName = sanitizeDriveName_([
      template.filePrefix,
      caseId,
      suffix ? suffix.replace(/^_/, '') : '',
      rowData.submission.name
    ].filter(function(part) { return Boolean(part); }).join('_'));
    const result = createDocumentFromTemplate_(template, fileName, caseFolder, data);
    result.caseId = caseId;
    result.rowNumber = rowData.rowNumber;
    if (!result.reused) appendDocumentRecord_(result);
    results.push(result);
  });
  return results;
}

function createDocumentFromTemplate_(template, fileName, destinationFolder, data) {
  const existingFiles = destinationFolder.getFilesByName(fileName);
  if (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    return {
      type: template.type,
      fileId: existingFile.getId(),
      name: existingFile.getName(),
      url: existingFile.getUrl(),
      pdfUrl: '',
      reused: true
    };
  }

  const templateFile = DriveApp.getFileById(template.templateId);
  const temporaryName = fileName + '__作成中_' + Utilities.getUuid().slice(0, 8);
  const copiedFile = templateFile.makeCopy(temporaryName, destinationFolder);
  try {
    const document = DocumentApp.openById(copiedFile.getId());
    replaceTemplateTags_(document, data);
    document.saveAndClose();
    copiedFile.setName(fileName);

    let pdfUrl = '';
    if (isPdfEnabled_()) {
      const pdfBlob = copiedFile.getAs(MimeType.PDF).setName(fileName + '.pdf');
      const pdfFile = destinationFolder.createFile(pdfBlob);
      pdfUrl = pdfFile.getUrl();
    }

    return {
      type: template.type,
      fileId: copiedFile.getId(),
      name: copiedFile.getName(),
      url: copiedFile.getUrl(),
      pdfUrl: pdfUrl,
      reused: false
    };
  } catch (error) {
    copiedFile.setTrashed(true);
    throw error;
  }
}

function replaceTemplateTags_(document, data) {
  const containers = [document.getBody()];
  const header = document.getHeader();
  const footer = document.getFooter();
  if (header) containers.push(header);
  if (footer) containers.push(footer);

  Object.keys(data).forEach(function(key) {
    const pattern = escapeRegExp_('{{' + key + '}}');
    // DocumentApp の replaceText は置換文字列をリテラル扱いするため、エスケープしない。
    const replacement = stringValue_(data[key]);
    containers.forEach(function(container) {
      container.replaceText(pattern, replacement);
    });
  });

  const unresolved = [];
  containers.forEach(function(container) {
    (container.getText().match(/\{\{[^{}]+\}\}/g) || []).forEach(function(tag) {
      unresolved.push(tag);
    });
  });
  if (unresolved.length) {
    throw new Error('未置換の差込タグがあります: ' + uniqueStrings_(unresolved).join(', '));
  }
}

function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
