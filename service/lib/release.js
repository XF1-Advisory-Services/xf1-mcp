export const repository = 'XF1-Advisory-Services/xf1-mcp';
export const templateId = 'xf1-base-six-sheet';
const hash = /^[a-f0-9]{64}$/i;
const version = /^\d+\.\d+\.\d+$/;

export function validateRelease(release) {
  if (!release || release.schemaVersion !== 1 || release.approval !== 'approved' ||
      release.templateId !== templateId || !version.test(release.version) ||
      release.generation !== 'from-scratch' || release.runner !== 'runner/New-XF1Workbook.ps1' ||
      !hash.test(release.manifestSha256) || !hash.test(release.package?.sha256) ||
      !Number.isSafeInteger(release.package?.bytes) || release.package.bytes < 1 || release.package.bytes > 5_000_000 ||
      release.bridge?.version !== '1.0.0' || !hash.test(release.bridge?.sha256)) {
    throw new Error('Approved release metadata is invalid. Stop; do not use a cached release.');
  }
  const base = `https://github.com/${repository}/releases/download/v${release.version}/`;
  if (release.package.url !== `${base}xf1-${release.version}.zip` ||
      release.bridge.url !== `${base}Invoke-XF1Build.ps1`) {
    throw new Error('Release assets must use the configured versioned GitHub release.');
  }
  return structuredClone(release);
}

export function resolveApprovedRelease(selection) {
  if (selection?.schemaVersion !== 1 || !selection.release) {
    throw new Error('No approved XF1 release is configured. Stop; do not use a cached release.');
  }
  return validateRelease(selection.release);
}
