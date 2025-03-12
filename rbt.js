const fs = require("fs");
const path = require("path");
const { Builder, By, Key, until } = require("selenium-webdriver");
const chrome = require("selenium-webdriver/chrome");

(async function main() {
  // 1) CONFIGURAÇÕES DE DOWNLOAD
  const downloadDir = path.join(process.cwd(), "downloads");
  if (!fs.existsSync(downloadDir)) {
    fs.mkdirSync(downloadDir, { recursive: true });
  }

  // Configurando as preferências do Chrome para downloads
  let options = new chrome.Options();
  options.setUserPreferences({
    "download.default_directory": downloadDir,
    "download.prompt_for_download": false,
    "download.directory_upgrade": true,
    "safebrowsing.enabled": true,
  });

  // Criação do driver com as opções definidas
  let driver = await new Builder()
    .forBrowser("chrome")
    .setChromeOptions(options)
    .build();

  try {
    // 2) LOGIN
    await driver.get("https://tosp-azr.ictsirio.com/tosp/");
    await driver.wait(
      until.elementLocated(By.xpath('//*[@id="username"]')),
      10000
    );
    await driver
      .findElement(By.xpath('//*[@id="username"]'))
      .sendKeys("alexandre.moura");
    await driver.sleep(1000);
    await driver
      .findElement(By.xpath('//*[@id="pass"]'))
      .sendKeys("486cf3", Key.RETURN);

    // 3) ACESSAR A URL DO DASHBOARD
    await driver.get(
      "https://tosp-azr.ictsirio.com/tosp/Workspace/load#/ManutenirMonitorOperacional"
    );
    await driver.wait(
      until.elementLocated(By.xpath('//*[@id="maincontent"]')),
      10000
    );

    // 4) CLICAR NO BOTÃO DA LUPA
    let lupaButton = await driver.wait(
      until.elementLocated(
        By.xpath(
          '//*[@id="maincontent"]/div[1]/div/div[1]/div/form/div/fieldset/table/tbody/tr/td[2]/button'
        )
      ),
      10000
    );
    await driver.wait(until.elementIsEnabled(lupaButton), 10000);
    await lupaButton.click();

    // 5) LOCALIZAR E CLICAR EM "JANELAS DE AGENDAMENTO"
    let container = await driver.wait(
      until.elementLocated(
        By.xpath('//*[contains(@id, "-grid-container")]/div[2]')
      ),
      20000
    );
    let found = false;
    for (let i = 0; i < 20; i++) {
      try {
        let janelasAgendamento = await container.findElement(
          By.xpath('.//div[contains(text(), "JANELAS DE AGENDAMENTO")]')
        );
        await driver.executeScript(
          "arguments[0].scrollIntoView({block: 'center'});",
          janelasAgendamento
        );
        await janelasAgendamento.click();
        found = true;
        break;
      } catch (err) {
        // Incrementa a rolagem do container
        await driver.executeScript("arguments[0].scrollTop += 300;", container);
      }
    }
    if (!found) {
      console.log(
        "Elemento 'JANELAS DE AGENDAMENTO' não encontrado após rolagem."
      );
      return;
    }

    // 6) CLICAR NO BOTÃO "SETAS"
    let setasButton = await driver.wait(
      until.elementLocated(
        By.xpath(
          '//*[@id="maincontent"]/div[1]/div/div[1]/div/div/fieldset/div[1]/button[1]'
        )
      ),
      10000
    );
    await driver.wait(until.elementIsEnabled(setasButton), 10000);
    await setasButton.click();

    // 7) ESPERAR O MODAL ABRIR
    const modalXpath = '//*[@id="visualizarMonitorOperacional"]/div/div';
    await driver.wait(until.elementLocated(By.xpath(modalXpath)), 30000);
    console.log("Modal do Power BI aberto!");

    // 8) LOCALIZAR O IFRAME DENTRO DO MODAL
    let reportIframe = await driver.wait(
      until.elementLocated(By.css("#report-container iframe")),
      20000
    );
    // 9) TROCAR PARA O CONTEXTO DO IFRAME
    await driver.switchTo().frame(reportIframe);
    console.log("Trocamos para o contexto do iframe do Power BI.");

    // 10) ESPERAR O POWER BI RENDERIZAR
    await driver.wait(until.elementLocated(By.id("pvExplorationHost")), 30000);
    console.log("Conteúdo do Power BI carregado dentro do iframe!");

    // 11) LOCALIZAR TODAS AS CÉLULAS DA TABELA
    const cellsXpath =
      '//*[@id="pvExplorationHost"]//visual-container-repeat/visual-container[7]//visual-modern//div/div/div[2]/div[1]/div[2]/div/div[*]/div/div/div[*]';
    let cells = await driver.wait(
      until.elementsLocated(By.xpath(cellsXpath)),
      20000
    );
    console.log(
      `Foram encontradas ${cells.length} células na tabela do Power BI.`
    );

    if (cells.length === 0) {
      console.log("Nenhuma célula encontrada na tabela.");
      return;
    }

    // 12) CLICAR NA PRIMEIRA CÉLULA PARA INICIAR A EXPORTAÇÃO
    let firstCell = cells[0];
    await driver.executeScript(
      "arguments[0].scrollIntoView({block: 'center'});",
      firstCell
    );
    await driver.sleep(500);
    await driver.actions().move({ origin: firstCell }).perform();
    await driver.sleep(500);
    await driver.executeScript("arguments[0].click();", firstCell);
    const cellText = await firstCell.getText();
    console.log(`Primeira célula clicada: ${cellText.trim()}`);
    await driver.sleep(1000);

    // 12.1) ROLAGEM "PASSO A PASSO" PARA CARREGAR TODA A TABELA
    try {
      let scrollContainer = await driver.wait(
        until.elementLocated(
          By.xpath(
            '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[7]/transform/div/div[3]/div/div/visual-modern/div/div/div[2]/div[4]/div'
          )
        ),
        10000
      );
      let lastHeight = 0;
      while (true) {
        await driver.executeScript(
          "arguments[0].scrollTop = arguments[0].scrollHeight;",
          scrollContainer
        );
        await driver.sleep(1000);
        let newHeight = await driver.executeScript(
          "return arguments[0].scrollTop;",
          scrollContainer
        );
        if (newHeight === lastHeight) break;
        lastHeight = newHeight;
      }
      console.log("Rolagem completa no container da barra de rolagem.");
    } catch (err) {
      console.log("Erro ao realizar a rolagem no elemento:", err);
      return;
    }
    await driver.sleep(2000);

    // 13) LOCALIZAR O BOTÃO "..." PELO XPATH FIXO
    const ellipsisButtonXpath =
      '//*[@id="pvExplorationHost"]/div/div/exploration/div/explore-canvas/div/div[2]/div/div[2]/div[2]/visual-container-repeat/visual-container[7]/transform/div/visual-container-header/div/div/div/visual-container-options-menu/visual-header-item-container/div/button';
    let ellipsisButton;
    try {
      ellipsisButton = await driver.findElement(By.xpath(ellipsisButtonXpath));
      if (await ellipsisButton.isDisplayed()) {
        console.log("Botão '...' localizado pelo XPath fixo.");
      } else {
        console.log("Botão '...' não está visível, mesmo com o XPath fixo.");
        return;
      }
    } catch (err) {
      console.log(
        "Não foi possível localizar o botão '...' usando o XPath fixo."
      );
      return;
    }
    await driver.sleep(2000);

    // 14) CLICAR NO BOTÃO "..."
    await driver.executeScript(
      "arguments[0].scrollIntoView(true);",
      ellipsisButton
    );
    await driver.sleep(500);
    await driver.executeScript("arguments[0].click();", ellipsisButton);
    console.log("Botão '...' clicado com sucesso!");

    // 14.1) CLICAR EM "EXPORTAR DADOS"
    const exportarDadosXpath = '//*[@id="0"]';
    try {
      let exportarDadosButton = await driver.wait(
        until.elementLocated(By.xpath(exportarDadosXpath)),
        20000
      );
      await driver.wait(until.elementIsEnabled(exportarDadosButton), 20000);
      await driver.executeScript("arguments[0].click();", exportarDadosButton);
      console.log("Opção 'Exportar Dados' clicada com sucesso!");
    } catch (err) {
      console.log("Erro ao clicar em 'Exportar Dados':", err);
      return;
    }

    // 15) CLICAR EM "DADOS RESUMIDOS"
    const dadosResumidosXpath =
      '//*[@id="pbi-radio-button-1"]/label/section/span';
    try {
      let dadosResumidosButton = await driver.wait(
        until.elementLocated(By.xpath(dadosResumidosXpath)),
        20000
      );
      await driver.wait(until.elementIsEnabled(dadosResumidosButton), 20000);
      await driver.executeScript("arguments[0].click();", dadosResumidosButton);
      console.log("Opção 'Dados resumidos' clicada com sucesso!");
    } catch (err) {
      console.log("Erro ao clicar em 'Dados resumidos':", err);
    }

    // 16) CLICAR NO BOTÃO "EXPORTAR"
    const exportarButtonXpath =
      '//*[@id="mat-mdc-dialog-0"]/div/div/export-data-dialog/mat-dialog-actions/button[1]';
    try {
      let exportarButton = await driver.wait(
        until.elementLocated(By.xpath(exportarButtonXpath)),
        20000
      );
      await driver.wait(until.elementIsEnabled(exportarButton), 20000);
      await driver.executeScript("arguments[0].click();", exportarButton);
      console.log("Botão 'Exportar' clicado com sucesso!");
    } catch (err) {
      console.log("Erro ao clicar no botão 'Exportar':", err);
    }

    // 17) AGUARDAR O DOWNLOAD DO ARQUIVO .XLSX
    const downloadedFile = path.join(downloadDir, "data.xlsx");
    const timeout = 60; // segundos
    let elapsed = 0;
    while (!fs.existsSync(downloadedFile) && elapsed < timeout) {
      await driver.sleep(1000);
      elapsed++;
    }
    if (fs.existsSync(downloadedFile)) {
      console.log("Download concluído:", downloadedFile);
    } else {
      console.log(
        `Timeout: arquivo data.xlsx não foi baixado em ${timeout} segundos.`
      );
    }
  } finally {
    await driver.quit();
  }
})().catch((err) => console.error(err));
