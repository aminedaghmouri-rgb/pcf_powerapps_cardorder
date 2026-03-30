import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class CardOrderAgent
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private context!: ComponentFramework.Context<IInputs>;
  private notifyOutputChanged!: () => void;
  private container!: HTMLDivElement;
  private inputElement!: HTMLInputElement;
  private currentValue: string = "";

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.container = container;

    this.currentValue = context.parameters.inputText.raw ?? "";

    const wrapper = document.createElement("div");
    wrapper.className = "card-order-agent";

    const title = document.createElement("div");
    title.className = "card-order-agent__title";
    title.textContent = "Card Order Agent";

    this.inputElement = document.createElement("input");
    this.inputElement.className = "card-order-agent__input";
    this.inputElement.type = "text";
    this.inputElement.placeholder = "Type a value";
    this.inputElement.value = this.currentValue;
    this.inputElement.addEventListener("input", this.onInputChanged);

    const hint = document.createElement("div");
    hint.className = "card-order-agent__hint";
    hint.textContent = "Output value is returned in upper-case.";

    wrapper.appendChild(title);
    wrapper.appendChild(this.inputElement);
    wrapper.appendChild(hint);

    this.container.appendChild(wrapper);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    const incomingValue = context.parameters.inputText.raw ?? "";

    if (incomingValue !== this.currentValue) {
      this.currentValue = incomingValue;
      if (this.inputElement) {
        this.inputElement.value = incomingValue;
      }
    }
  }

  public getOutputs(): IOutputs {
    return {
      processedText: this.currentValue.trim().toUpperCase()
    };
  }

  public destroy(): void {
    if (this.inputElement) {
      this.inputElement.removeEventListener("input", this.onInputChanged);
    }
  }

  private onInputChanged = (event: Event): void => {
    const target = event.target as HTMLInputElement;
    this.currentValue = target.value;
    this.notifyOutputChanged();
  };
}
