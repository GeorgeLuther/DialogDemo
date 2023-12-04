import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";
import ColorPickerDialog from "./ColorPickerDialog";
import { IColor } from "office-ui-fabric-react/lib/Color";
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDialogDemoCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = "DialogDemoCommandSet";

export default class DialogDemoCommandSet extends BaseListViewCommandSet<IDialogDemoCommandSetProperties> {
  private _colorCode: IColor;
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized DialogDemoCommandSet");

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "COMMAND_1":
        Dialog.alert(`${this.properties.sampleTextOne}`);
        const dialog: ColorPickerDialog = new ColorPickerDialog();
        dialog.message = "Pick a color:";
        // Use 'FFFFFF' as the default color for first usage
        let defaultColor: IColor = {
          hex: "FFFFFF",
          str: "",
          r: null,
          g: null,
          b: null,
          h: null,
          s: null,
          v: null,
        };
        dialog.colorCode = this._colorCode || defaultColor;
        dialog.show().then(() => {
          this._colorCode = dialog.colorCode;
          Dialog.alert(`Picked color: ${dialog.colorCode.hex}`);
        });
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _onListViewStateChanged = (
    args: ListViewStateChangedEventArgs
  ): void => {
    Log.info(LOG_SOURCE, "List view state changed");

    const compareOneCommand: Command = this.tryGetCommand("COMMAND_1");
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible =
        this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  };
}
