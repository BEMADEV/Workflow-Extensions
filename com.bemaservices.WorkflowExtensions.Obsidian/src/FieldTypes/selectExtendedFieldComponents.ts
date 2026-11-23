import { computed, defineComponent, ref, useAttrs, watch } from "vue";
import CheckBoxList from "@Obsidian/Controls/checkBoxList.obs";
import DropDownList from "@Obsidian/Controls/dropDownList.obs";
import ListBox from "@Obsidian/Controls/listBox.obs";
import NumberBox from "@Obsidian/Controls/numberBox.obs";
import RadioButtonList from "@Obsidian/Controls/radioButtonList.obs";
import TextBox from "@Obsidian/Controls/textBox.obs";
import { toNumberOrNull } from "@Obsidian/Utility/numberUtils";
import { ListItemBag } from "@Obsidian/ViewModels/Utility/listItemBag";
import {
    ConfigurationValueKey,
    getChildOptions,
    getParentOptions
} from "./selectExtendedField.partial";
import { getFieldConfigurationProps, getFieldEditorProps } from "@Obsidian/FieldTypes/utils";

export const EditComponent = defineComponent({
    name: "SelectExtendedField.Edit",

    components: {
        CheckBoxList,
        DropDownList,
        ListBox,
        RadioButtonList
    },

    inheritAttrs: false,

    props: getFieldEditorProps(),

    emits: [
        "update:modelValue"
    ],

    setup(props, { emit }) {
        const attrs = useAttrs();
        const parentValue = ref("");
        const childValues = ref<string[]>([]);
        let isSynchronizing = false;

        const fieldLabel = computed(() => typeof attrs.label === "string" ? attrs.label : undefined);
        const fieldHelp = computed(() => typeof attrs.help === "string" ? attrs.help : undefined);
        const validationRules = computed(() => typeof attrs.rules === "string" ? attrs.rules : undefined);

        const parentOptions = computed((): ListItemBag[] => getParentOptions(props.configurationValues));
        const childOptions = computed((): ListItemBag[] => {
            return getChildOptions(props.configurationValues)
                .filter(option => (option.parentValue ?? "") === parentValue.value)
                .map(option => ({
                    value: option.value,
                    text: option.text
                }));
        });

        const parentFieldType = computed(() => props.configurationValues[ConfigurationValueKey.ParentFieldType] ?? "ddl");
        const childFieldType = computed(() => props.configurationValues[ConfigurationValueKey.ChildFieldType] ?? "ddl");
        const parentRepeatColumns = computed(() => toNumberOrNull(props.configurationValues[ConfigurationValueKey.ParentRepeatColumns]) ?? 0);
        const childRepeatColumns = computed(() => toNumberOrNull(props.configurationValues[ConfigurationValueKey.ChildRepeatColumns]) ?? 0);
        const isParentRadio = computed(() => parentFieldType.value === "rb");
        const isParentEnhanced = computed(() => parentFieldType.value === "ddl_single_enhanced");
        const isChildRadio = computed(() => childFieldType.value === "rb");
        const isChildCheckBox = computed(() => childFieldType.value === "cb");
        const isChildMultiSelect = computed(() => childFieldType.value === "ddl_multi_enhanced");
        const isChildEnhanced = computed(() => childFieldType.value === "ddl_single_enhanced");

        const parentRadioOptions = computed((): ListItemBag[] => [
            {
                value: "",
                text: "None"
            },
            ...parentOptions.value
        ]);

        const singleChildValue = computed({
            get: (): string => childValues.value[0] ?? "",
            set: (value: string): void => {
                childValues.value = value ? [value] : [];
            }
        });

        const childRadioOptions = computed((): ListItemBag[] => [
            {
                value: "",
                text: "None"
            },
            ...childOptions.value
        ]);

        const emitValue = (): void => {
            const childValue = childValues.value.join(",");
            const value = parentValue.value || childValue
                ? `${parentValue.value}|${childValue}`
                : "";

            if (value !== props.modelValue) {
                emit("update:modelValue", value);
            }
        };

        watch(() => props.modelValue, value => {
            isSynchronizing = true;

            const [newParentValue = "", newChildValue = ""] = (value ?? "").split("|", 2);
            parentValue.value = newParentValue;
            childValues.value = newChildValue.split(",").filter(childValue => childValue !== "");

            isSynchronizing = false;
        }, {
            immediate: true,
            flush: "sync"
        });

        watch(parentValue, () => {
            if (isSynchronizing) {
                return;
            }

            isSynchronizing = true;
            childValues.value = [];
            isSynchronizing = false;
            emitValue();
        }, {
            flush: "sync"
        });

        watch(childValues, () => {
            if (!isSynchronizing) {
                emitValue();
            }
        }, {
            deep: true,
            flush: "sync"
        });

        return {
            childOptions,
            childRadioOptions,
            childRepeatColumns,
            childValues,
            fieldHelp,
            fieldLabel,
            isChildCheckBox,
            isChildEnhanced,
            isChildMultiSelect,
            isChildRadio,
            isParentEnhanced,
            isParentRadio,
            parentOptions,
            parentRadioOptions,
            parentRepeatColumns,
            parentValue,
            singleChildValue,
            validationRules
        };
    },

    template: `
<div>
    <RadioButtonList v-if="isParentRadio"
        v-model="parentValue"
        :items="parentRadioOptions"
        :label="fieldLabel"
        :help="fieldHelp"
        :rules="validationRules"
        :repeatColumns="parentRepeatColumns"
        horizontal />
    <DropDownList v-else
        v-model="parentValue"
        :items="parentOptions"
        :label="fieldLabel"
        :help="fieldHelp"
        :rules="validationRules"
        :enhanceForLongLists="isParentEnhanced" />

    <template v-if="parentValue">
        <RadioButtonList v-if="isChildRadio"
            v-model="singleChildValue"
            :items="childRadioOptions"
            :rules="validationRules"
            :repeatColumns="childRepeatColumns"
            horizontal />
        <CheckBoxList v-else-if="isChildCheckBox"
            v-model="childValues"
            :items="childOptions"
            :rules="validationRules"
            :repeatColumns="childRepeatColumns"
            horizontal />
        <ListBox v-else-if="isChildMultiSelect"
            v-model="childValues"
            :items="childOptions"
            :rules="validationRules"
            enhanceForLongLists />
        <DropDownList v-else
            v-model="singleChildValue"
            :items="childOptions"
            :rules="validationRules"
            :enhanceForLongLists="isChildEnhanced" />
    </template>
</div>
`
});

const parentControlTypeOptions: ListItemBag[] = [
    {
        value: "ddl",
        text: "Drop Down List"
    },
    {
        value: "ddl_single_enhanced",
        text: "Single Select Drop Down List (Enhanced for Long Lists)"
    },
    {
        value: "rb",
        text: "Radio Buttons"
    }
];

const childControlTypeOptions: ListItemBag[] = [
    ...parentControlTypeOptions,
    {
        value: "ddl_multi_enhanced",
        text: "Multi Select Drop Down List (Enhanced for Long Lists)"
    },
    {
        value: "cb",
        text: "Check Boxes"
    }
];

export const ConfigurationComponent = defineComponent({
    name: "SelectExtendedField.Configuration",

    components: {
        DropDownList,
        NumberBox,
        TextBox
    },

    props: getFieldConfigurationProps(),

    emits: [
        "update:modelValue",
        "updateConfiguration",
        "updateConfigurationValue"
    ],

    setup(props, { emit }) {
        const parentRawValues = ref("");
        const internalParentRawValues = ref("");
        const parentControlType = ref("ddl");
        const parentRepeatColumns = ref<number | null>(null);
        const childRawValues = ref("");
        const internalChildRawValues = ref("");
        const childControlType = ref("ddl");
        const childRepeatColumns = ref<number | null>(null);

        const isParentRadio = computed(() => parentControlType.value === "rb");
        const isChildRepeatedList = computed(() => childControlType.value === "rb" || childControlType.value === "cb");

        const maybeUpdateModelValue = (): boolean => {
            const newValue: Record<string, string> = { ...props.modelValue };

            newValue[ConfigurationValueKey.CustomParentValues] = internalParentRawValues.value ?? "";
            newValue[ConfigurationValueKey.ParentFieldType] = parentControlType.value ?? "ddl";
            newValue[ConfigurationValueKey.ParentRepeatColumns] = parentRepeatColumns.value?.toString() ?? "";
            newValue[ConfigurationValueKey.CustomChildValues] = internalChildRawValues.value ?? "";
            newValue[ConfigurationValueKey.ChildFieldType] = childControlType.value ?? "ddl";
            newValue[ConfigurationValueKey.ChildRepeatColumns] = childRepeatColumns.value?.toString() ?? "";

            const configurationKeys = [
                ConfigurationValueKey.CustomParentValues,
                ConfigurationValueKey.ParentFieldType,
                ConfigurationValueKey.ParentRepeatColumns,
                ConfigurationValueKey.CustomChildValues,
                ConfigurationValueKey.ChildFieldType,
                ConfigurationValueKey.ChildRepeatColumns
            ];

            const hasChanged = configurationKeys.some(key => newValue[key] !== (props.modelValue[key] ?? ""));
            if (!hasChanged) {
                return false;
            }

            emit("update:modelValue", newValue);
            return true;
        };

        const maybeUpdateConfigurationValue = (key: string, value: string): void => {
            if (maybeUpdateModelValue()) {
                emit("updateConfigurationValue", key, value);
            }
        };

        const onParentValuesBlur = (): void => {
            internalParentRawValues.value = parentRawValues.value;
        };

        const onChildValuesBlur = (): void => {
            internalChildRawValues.value = childRawValues.value;
        };

        watch(() => [props.modelValue, props.configurationProperties], () => {
            parentRawValues.value = props.modelValue[ConfigurationValueKey.CustomParentValues] ?? "";
            internalParentRawValues.value = parentRawValues.value;
            parentControlType.value = props.modelValue[ConfigurationValueKey.ParentFieldType] ?? "ddl";
            parentRepeatColumns.value = toNumberOrNull(props.modelValue[ConfigurationValueKey.ParentRepeatColumns]);
            childRawValues.value = props.modelValue[ConfigurationValueKey.CustomChildValues] ?? "";
            internalChildRawValues.value = childRawValues.value;
            childControlType.value = props.modelValue[ConfigurationValueKey.ChildFieldType] ?? "ddl";
            childRepeatColumns.value = toNumberOrNull(props.modelValue[ConfigurationValueKey.ChildRepeatColumns]);
        }, {
            immediate: true
        });

        watch([internalParentRawValues, internalChildRawValues], () => {
            if (maybeUpdateModelValue()) {
                emit("updateConfiguration");
            }
        });

        watch(parentControlType, value => maybeUpdateConfigurationValue(ConfigurationValueKey.ParentFieldType, value ?? "ddl"));
        watch(parentRepeatColumns, value => maybeUpdateConfigurationValue(ConfigurationValueKey.ParentRepeatColumns, value?.toString() ?? ""));
        watch(childControlType, value => maybeUpdateConfigurationValue(ConfigurationValueKey.ChildFieldType, value ?? "ddl"));
        watch(childRepeatColumns, value => maybeUpdateConfigurationValue(ConfigurationValueKey.ChildRepeatColumns, value?.toString() ?? ""));

        return {
            childControlType,
            childControlTypeOptions,
            childRawValues,
            childRepeatColumns,
            isChildRepeatedList,
            isParentRadio,
            onChildValuesBlur,
            onParentValuesBlur,
            parentControlType,
            parentControlTypeOptions,
            parentRawValues,
            parentRepeatColumns
        };
    },

    template: `
<div>
    <TextBox v-model="parentRawValues"
        label="Parent Values"
        help="The source of the values to display in the parent list. Format is either 'value1,value2,value3,...', 'value1^text1,value2^text2,value3^text3,...', or a SQL Select statement that returns 'Value' and 'Text' columns <span class='tip tip-lava'></span>."
        textMode="multiline"
        @blur="onParentValuesBlur" />

    <DropDownList v-model="parentControlType"
        label="Parent Control Type"
        help="The type of control to use for selecting a parent value."
        :items="parentControlTypeOptions"
        :showBlankItem="false" />

    <NumberBox v-if="isParentRadio"
        v-model="parentRepeatColumns"
        label="Parent Columns"
        help="The number of columns used to display the parent list. Leave blank or use 0 for the default." />

    <TextBox v-model="childRawValues"
        label="Child Values"
        help="The source of the child values. Format is 'value1^parent1,value2^parent2,...', 'value1^parent1^text1,value2^parent2^text2,...', or a SQL Select statement that returns 'Value', 'ParentValue', and 'Text' columns <span class='tip tip-lava'></span>."
        textMode="multiline"
        @blur="onChildValuesBlur" />

    <DropDownList v-model="childControlType"
        label="Child Control Type"
        help="The type of control to use for selecting one or more child values."
        :items="childControlTypeOptions"
        :showBlankItem="false" />

    <NumberBox v-if="isChildRepeatedList"
        v-model="childRepeatColumns"
        label="Child Columns"
        help="The number of columns used to display the child list. Leave blank or use 0 for the default." />
</div>
`
});