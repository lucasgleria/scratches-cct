
import asyncio
from playwright.async_api import async_playwright, expect

def read_file_content(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        return f.read()

async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        try:
            # A more robust mock that better emulates the chaining behavior of google.script.run
            mock_script = """
                window.google = {
                    script: {
                        run: new Proxy({}, {
                            get(target, prop) {
                                const runner = {
                                    _success: () => {},
                                    _failure: () => {},
                                    withSuccessHandler(handler) {
                                        this._success = handler;
                                        return this;
                                    },
                                    withFailureHandler(handler) {
                                        this._failure = handler;
                                        return this;
                                    },
                                };

                                // This function will be the one that is called e.g. google.script.run.getSpreadsheets()
                                runner[prop] = (...args) => {
                                    if (prop === 'getSpreadsheets') {
                                        runner._success({ ok: true, payload: [{id: '1Qrzq3NatjRtLE8CiQbMiWRHvwFgUKA5ymimoR6JAsV0', name: 'CCT Teste'}] });
                                    } else if (prop === 'getAllHouses') {
                                        runner._success({ ok: true, payload: ['HOUSE1', 'HOUSE2', 'HOUSE3'] });
                                    } else if (prop === 'getDataForHouse') {
                                        runner._success({
                                            ok: true,
                                            payload: {
                                                mawb: '123-45678901', house: args[1], refs: ['REF1'], consignees: ['CONSIGNEE1'],
                                                entregas: ['ENTREGA1'], dtas: ['DTA1'], previsoes: ['PREVISAO1'],
                                                responsaveis: [], observacoes: []
                                            }
                                        });
                                    } else if (prop === 'saveEntries') {
                                        const payload = args[0];
                                        runner._success({ ok: true, payload: { inserted: payload.houses.length, updated: 0, removed_duplicates: 0 } });
                                    }
                                };
                                return runner[prop] || runner;
                            }
                        })
                    }
                };
            """
            await page.add_init_script(mock_script)

            # Manually build the full HTML by replacing includes
            importacao_html = read_file_content('importacao.html')
            style_css = read_file_content('style.css.html')
            editor_modal_html = read_file_content('editor-modal.html')
            servicos_modal_html = read_file_content('servicos-modal.html')
            main_js = read_file_content('main.js.html')

            final_html = importacao_html.replace("<?!= include('style.css'); ?>", f"<style>{style_css}</style>")
            final_html = final_html.replace("<?!= include('editor-modal.html'); ?>", editor_modal_html)
            final_html = final_html.replace("<?!= include('servicos-modal.html'); ?>", servicos_modal_html)
            final_html = final_html.replace("<?!= include('main.js'); ?>", main_js)

            await page.set_content(final_html, wait_until="networkidle")

            # 1. Select a spreadsheet
            await page.select_option('#spreadsheet', label='CCT Teste')

            # 2. Click the "Serviços" toggle (targeting the visible slider)
            await page.locator('label[for="servicosToggle"] .slider').click()

            # 3. Wait for the modal and click a HOUSE
            await expect(page.locator('#servicos-modal')).to_be_visible()
            await page.get_by_role('list').locator('li.house-list-item').filter(has_text='HOUSE1').click()

            # 4. Add a "Responsável"
            await expect(page.get_by_placeholder('Digite um responsável')).to_be_visible()
            await page.get_by_placeholder('Digite um responsável').fill('Jules')

            # 5. Click "Salvar"
            await page.get_by_role('button', name='Salvar Dados').click()

            # 6. Verify success alert
            await expect(page.locator('.alert-success')).to_be_visible(timeout=5000)
            await expect(page.locator('.alert-success')).to_contain_text('Dados salvos com sucesso!')

            # 7. Take a screenshot
            await page.screenshot(path="/home/jules/verification.png")
            print("Screenshot taken successfully.")

        finally:
            await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
