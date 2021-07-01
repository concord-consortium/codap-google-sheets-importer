import { OptionTypeBase, GroupTypeBase, StylesConfig } from "react-select";

export const customStyles: StylesConfig<
  OptionTypeBase,
  true,
  GroupTypeBase<OptionTypeBase>
> = {
  control: (provided, state) => ({
    ...provided,
    borderColor: "rgba(23, 121, 145, 0.75)",
    width: "300px",
  }),
};
