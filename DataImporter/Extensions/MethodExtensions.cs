using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;

namespace DataImporter.Extensions
{
    public static class MethodExtensions
    {
        public static void AsRead(this Methods methods, string name, string displayName, Validation validation, MethodParameters parameters, InputProperties inputs, ReturnProperties returns)
        {
            Create(methods, name, displayName, MethodType.Read, validation, parameters, inputs, returns);
        }

        private static void Create(this Methods methods, string name, string displayName, MethodType methodType, Validation validation, MethodParameters parameters, InputProperties inputs, ReturnProperties returns)
        {
            var meta = new MetaData
            {
                DisplayName = displayName,
                Description = ""
            };

            Create(methods, name, methodType, meta, validation, parameters, inputs, returns);
        }

        private static void Create(this Methods methods, string name, MethodType methodType, MetaData metaData, Validation validation, MethodParameters parameters, InputProperties inputs, ReturnProperties returns)
        {
            if (methods.Contains(name)) return;
            
            methods.Create(new Method
            {
                Name = name,
                Type = methodType,
                MetaData = metaData ?? new MetaData(),
                Validation = validation ?? new Validation(),
                MethodParameters = parameters ?? new MethodParameters(),
                InputProperties = inputs ?? new InputProperties(),
                ReturnProperties = returns ?? new ReturnProperties()
            });
        }
    }
}