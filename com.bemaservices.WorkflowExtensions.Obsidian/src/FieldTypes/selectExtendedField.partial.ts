import { Component } from "vue";
import { defineAsyncComponent } from "@Obsidian/Utility/component";
import { FieldTypeBase } from "@Obsidian/FieldTypes/fieldType";
import { ListItemBag } from "@Obsidian/ViewModels/Utility/listItemBag";

export const enum ConfigurationValueKey {
    ParentValues = "parent_values",
    ParentFieldType = "parent_fieldtype",
    ParentRepeatColumns = "parent_repeatColumns",
    ChildValues = "child_values",
    ChildFieldType = "child_fieldtype",
    ChildRepeatColumns = "child_repeatColumns",

    /** Only used while editing the field type configuration. */
    CustomParentValues = "customParentValues",
    CustomChildValues = "customChildValues"
}

export type SelectExtendedOption = ListItemBag & {
    parentValue?: string;
};

export function getParentOptions(configurationValues: Record<string, string>): ListItemBag[] {
    try {
        const options = JSON.parse(configurationValues[ConfigurationValueKey.ParentValues] ?? "[]") as ListItemBag[];
        return Array.isArray(options) ? options : [];
    }
    catch {
        return [];
    }
}

export function getChildOptions(configurationValues: Record<string, string>): SelectExtendedOption[] {
    try {
        const options = JSON.parse(configurationValues[ConfigurationValueKey.ChildValues] ?? "[]") as SelectExtendedOption[];
        return Array.isArray(options) ? options : [];
    }
    catch {
        return [];
    }
}

const editComponent = defineAsyncComponent(async () => {
    return (await import("./selectExtendedFieldComponents")).EditComponent;
});

const configurationComponent = defineAsyncComponent(async () => {
    return (await import("./selectExtendedFieldComponents")).ConfigurationComponent;
});

export class SelectExtendedFieldType extends FieldTypeBase {
    public override getTextValue(value: string, configurationValues: Record<string, string>): string {
        if (!value) {
            return "";
        }

        const [parentValue = "", childValue = ""] = value.split("|", 2);
        const selectedParentValues = parentValue.split(",").filter(v => v !== "");
        const selectedChildValues = childValue.split(",").filter(v => v !== "");

        const parentText = getParentOptions(configurationValues)
            .filter(option => selectedParentValues.includes(option.value ?? ""))
            .map(option => option.text ?? "")
            .join(", ");

        const childText = getChildOptions(configurationValues)
            .filter(option => selectedChildValues.includes(option.value ?? ""))
            .map(option => option.text ?? "")
            .join(", ");

        return parentText || childText ? `${parentText}: ${childText}` : "";
    }

    public override getEditComponent(): Component {
        return editComponent;
    }

    public override getConfigurationComponent(): Component {
        return configurationComponent;
    }

    public override isFilterable(): boolean {
        return false;
    }
}