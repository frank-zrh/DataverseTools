using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDataverseTool
{
    public static class ActionType
    {
        public static string[] actions = { "Create", "Append", "Update", "Clear" };

        public static ActionTypeEnum GetActionType(string action)
        {
            switch (action.ToLower())
            {
                case "create":
                    return ActionTypeEnum.Create;
                case "append":
                    return ActionTypeEnum.Append;
                case "update":
                    return ActionTypeEnum.Update;
                case "clear":
                    return ActionTypeEnum.Clear;
                default:
                    return ActionTypeEnum.Update;
            }
        }

        public static string GetActionName(ActionTypeEnum action)
        {
            switch (action)
            {
                case ActionTypeEnum.Create:
                    return "Create";
                case ActionTypeEnum.Append:
                    return "Append";
                case ActionTypeEnum.Update:
                    return "Update";
                case ActionTypeEnum.Clear:
                    return "Clear";
                default:
                    return "Update";
            }
        }
    }

    public enum ActionTypeEnum
    {
        Create,
        Append,
        Update,
        Clear
    }

}
