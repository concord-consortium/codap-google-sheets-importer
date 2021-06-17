import { useState, useCallback } from "react";

interface ElementWithValue {
  value: string;
}

export function useInput<T, E extends ElementWithValue>(
  initialValue: T | string,
  extraAction?: (newValue: T | string) => void
): [
  T | string,
  (e: React.ChangeEvent<E>) => void,
  React.Dispatch<React.SetStateAction<T | string>>
] {
  const [inputValue, setInputValue] = useState<T | string>(initialValue);
  const onChange = useCallback(
    (event: React.ChangeEvent<E>) => {
      setInputValue(event.target.value);
      if (extraAction !== undefined) {
        extraAction(event.target.value);
      }
    },
    [setInputValue, extraAction]
  );
  return [inputValue, onChange, setInputValue];
}
