import { test, expect, Page } from '@playwright/test';

const APP_URL = 'https://islamar-reservas.streamlit.app/';
const USER    = process.env.ISLAMAR_USER || 'festeban';
const PASS    = process.env.ISLAMAR_PASS ?? '';

function app(page: Page) {
  return page.frameLocator('iframe[title="streamlitApp"]');
}

async function login(page: Page) {
  await page.goto(APP_URL);
  const frame = app(page);
  await frame.getByRole('textbox', { name: 'Usuario' }).fill(USER);
  await frame.getByRole('textbox', { name: 'Contraseña' }).fill(PASS);
  await frame.getByTestId('stBaseButton-secondaryFormSubmit').click();
  await expect(frame.getByText('NAVEGACIÓN')).toBeVisible({ timeout: 20000 });
}

test('smoke: login y navegación por el sidebar', async ({ page }) => {
  test.skip(!PASS, 'Define ISLAMAR_PASS en el entorno antes de correr.');

  await login(page);
  const frame = app(page);

  await frame.getByText('📋 Listado Raquel').click();
  await expect(frame.getByText(/Listado.*Raquel/i)).toBeVisible();

  await frame.getByText('🔒 Auditoría').click();
  await expect(frame.getByText(/Auditor[ií]a/)).toBeVisible();

  await frame.getByText('📅 Plantilla mensual').click();
  await expect(frame.getByText(/Plantilla mensual/i)).toBeVisible();
});
