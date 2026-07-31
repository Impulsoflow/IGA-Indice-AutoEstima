/*
 * IGA-Indice-AutoEstima — correção do recebimento legado
 *
 * Adicionar ao final do Código.gs do projeto "IGA AutoEstima".
 * Esta declaração substitui a função legado anterior, mantendo o DISC v1
 * compatível e registrando o resultado autoestima_v1 em Google Sheets.
 */
function handleLegacyTest(dados) {
  try {
    const isAutoestima = dados.testId === "autoestima_v1";
    let pdfBlob = null;
    let fileUrl = "";

    if (dados.pdfBase64 && dados.pdfFileName) {
      try {
        const pasta = DriveApp.getFolderById(CONFIG.PASTA_PDF_ID);
        const decoded = Utilities.base64Decode(dados.pdfBase64);
        pdfBlob = Utilities.newBlob(decoded, "application/pdf", dados.pdfFileName);
        fileUrl = pasta.createFile(pdfBlob).getUrl();
      } catch (erroPdf) {
        logEvent(isAutoestima ? "AUTOESTIMA_PDF_ERROR" : "LEGACY_PDF_ERROR", erroPdf.message);
        if (!isAutoestima) throw erroPdf;
      }
    }

    let nomeTeste = "Avaliação Instituto Impulso";
    let resumoTexto = dados.resumo || "";
    if (dados.testId === "disc_v1") {
      nomeTeste = "Pesquisa Impulso DISC";
      resumoTexto = `Perfil Primário: ${dados.perfilPrimario} (${dados.pctD}% D, ${dados.pctI}% I, ${dados.pctS}% S, ${dados.pctC}% C)\n` +
        `Perfil Secundário: ${dados.perfilSecundario}`;
    } else if (isAutoestima) {
      nomeTeste = "Avaliação Nível de Autoestima";
    }

    let sheetRow = null;
    if (isAutoestima) {
      const sheet = getOrCreateSheet("IGA_Autoestima", [
        "Timestamp", "Nome", "Email", "WhatsApp", "Idade", "Profissão", "Escolaridade",
        "IGA (%)", "Nível", "Perfil", "Descrição do Perfil", "Mínimo Saudável (%)",
        "Módulos (JSON)", "Resumo", "Link do PDF", "Arquivo PDF", "Status do E-mail"
      ]);

      sheet.appendRow([
        new Date().toISOString(),
        dados.nome || "",
        dados.email || "",
        dados.whatsapp || "",
        dados.idade || "",
        dados.profissao || "",
        dados.escolaridade || "",
        Number(dados.igaPct) || 0,
        dados.nivel || "",
        dados.perfilTitulo || "",
        dados.perfilTexto || "",
        Number(dados.minimoSaudavel) || 0,
        dados.modulosJson || "[]",
        resumoTexto,
        fileUrl,
        dados.pdfFileName || "",
        dados.sendEmail === true ? "PENDENTE" : "NÃO SOLICITADO"
      ]);
      sheetRow = sheet.getLastRow();
      logEvent("AUTOESTIMA_SAVED", `linha=${sheetRow} | ${dados.email || "sem e-mail"}`);
    }

    let emailOk = true;
    let emailError = "";
    if (dados.sendEmail === true) {
      try {
        const subject = `Resultado: ${nomeTeste} - ${dados.nome || "Participante"}`;
        const bodyCliente = `Olá, ${dados.nome || ""}!\n\nAgradecemos por realizar a ${nomeTeste}.\n\nSegue em anexo o PDF com o seu resultado completo.\n\nResumo:\n${resumoTexto}\n\nUm abraço,\nRogério Braga\nInstituto Impulso Coaching de Liderança`;
        const bodyAdmin = `Novo resultado!\nTeste: ${nomeTeste}\nNome: ${dados.nome || ""}\nE-mail: ${dados.email || ""}\n\nResumo:\n${resumoTexto}\n\nDrive: ${fileUrl}`;
        const anexos = pdfBlob ? [pdfBlob] : [];

        if (dados.email && String(dados.email).includes("@")) {
          MailApp.sendEmail({ to: dados.email, subject, body: bodyCliente, attachments: anexos });
        }
        MailApp.sendEmail({ to: CONFIG.ADMIN_EMAIL, subject: `[NOVO] ${subject}`, body: bodyAdmin, attachments: anexos });
      } catch (erroEmail) {
        emailOk = false;
        emailError = erroEmail.message;
        logEvent(isAutoestima ? "AUTOESTIMA_EMAIL_ERROR" : "LEGACY_EMAIL_ERROR", erroEmail.message);
      }
    }

    if (isAutoestima && sheetRow) {
      const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName("IGA_Autoestima");
      sheet.getRange(sheetRow, 17).setValue(
        dados.sendEmail === true ? (emailOk ? "ENVIADO" : "ERRO: " + emailError) : "NÃO SOLICITADO"
      );
      SpreadsheetApp.flush();
    }

    return jsonOut({
      ok: true,
      status: "sucesso",
      testId: dados.testId || "",
      sheetRow,
      driveUrl: fileUrl,
      emailOk,
      emailError,
      message: isAutoestima ? "Resultado IGA registrado na planilha" : "Resultado legado processado"
    });
  } catch (erro) {
    logEvent("ERROR_LEGACY", erro.message);
    return jsonOut({ ok: false, error: erro.message });
  }
}
