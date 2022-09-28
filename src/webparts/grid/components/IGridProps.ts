import { ISelectedField } from "../GridWebPart";
import { ISharePointService } from "../services/spService";

export interface IGridProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  list: string;
  fields: ISelectedField[];
  fullWidth: string;
  orderBy: string;
  service: ISharePointService;
  title: string;
  footer: string;
  collapsed: boolean;
}
