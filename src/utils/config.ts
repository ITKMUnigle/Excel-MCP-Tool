export function parseBooleanEnv(value: string | undefined): boolean {
  if (!value) {
    return false;
  }

  const normalizedValue = value.trim().toLowerCase();
  return ['true', 'yes', 'y'].includes(normalizedValue);
}

export function isWriteAllowed(): boolean {
  return parseBooleanEnv(process.env.EXCEL_ALLOW_WRITE);
}
