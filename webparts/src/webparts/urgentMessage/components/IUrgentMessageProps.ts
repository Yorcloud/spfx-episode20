import {
  IReadonlyTheme,
} from "@microsoft/sp-component-base";

export interface IUrgentMessageProps {
  description: string;
  currentUser: string;
  list: string;
  label: string;
  message: string;
  themeVariant: IReadonlyTheme | undefined;
}
