import { registerFieldType } from "@Obsidian/Utility/fieldTypes";
import { SelectExtendedFieldType } from "./FieldTypes/selectExtendedField.partial";

const selectExtendedFieldTypeGuid = "e38f21d9-12ad-43bb-a6ba-e51b889c168d";

registerFieldType(selectExtendedFieldTypeGuid, new SelectExtendedFieldType());