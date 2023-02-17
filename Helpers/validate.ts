function validate<T>(obj: T, validations: string[]) {
  return validations.every(
    (key) =>
      ![undefined, null].includes(
        key.split(".").reduce((acc, cur) => acc?.[cur], obj)
      )
  );
}
