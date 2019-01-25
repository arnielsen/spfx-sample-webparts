import { IWebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IDelveBlogPostsProps {
  rowLimit: number;
  commandBar: boolean;
  context: IWebPartContext;
  people: IPropertyFieldGroupOrPerson[];
}
