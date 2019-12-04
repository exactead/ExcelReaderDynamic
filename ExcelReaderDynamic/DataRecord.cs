using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;

namespace ExcelReaderDynamic
{
    internal class DataRecord : DynamicObject
    {
        private readonly IDictionary<string, object> dictionary;

        /// <summary>
        /// DataRecordコンストラクタ（引数：IDictionary型）
        /// </summary>
        /// <param name="dictionary"></param>
        public DataRecord(IDictionary<string, object> dictionary) => this.dictionary = dictionary;

        /// <summary>
        /// DataRecordコンストラクタ（引数：匿名型(Anonymous object)）
        /// </summary>
        /// <param name="_object"></param>
        public DataRecord(object _object)
        {
            this.dictionary = _object.GetType()
                    .GetProperties()
                    .Where(x => x.CanRead)
                    .ToDictionary(v => v.Name, v => v.GetValue(_object));
        }

        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            // キーがないならNG
            if (!IsTypeCheck(binder.Name, value)) return false;
            dictionary[binder.Name] = value;
            return true;
        }

        public override bool TrySetIndex(SetIndexBinder binder, object[] indexes, object value)
        {
            var index = indexes[0];
            if (index is string)
            {
                if (IsTypeCheck((string)index, value))
                {
                    dictionary[(string)index] = value;
                    return true;
                }
            }
            if (index is int)
            {
                if (IsTypeCheck((int)index, value))
                {
                    dictionary[dictionary.ElementAt((int)index).Key] = value;
                    return true;
                }
            }
            return false;
        }

        private bool IsTypeCheck(string key, object value)
        {
            // 文字列で検索する場合はTryGetValueで存在するか確認
            if (!dictionary.TryGetValue(key, out var result)) return false;

            // 型が一致しない場合はNG
            return IsTypeMatch(result.GetType(), value.GetType());
        }

        private bool IsTypeCheck(int key, object value)
        {
            // 数値で検索する場合はElementAtOrDefaultでNULLか確認
            var baseValue = dictionary.ElementAtOrDefault(key).Value;
            if (baseValue == null) return false;

            // 型が一致しない場合はNG
            return IsTypeMatch(baseValue.GetType(), value.GetType());
        }

        private bool IsTypeMatch(Type baseType, Type valueType)
            => valueType.Equals(baseType) || valueType.IsSubclassOf(baseType);

        public override bool TryGetMember(GetMemberBinder binder, out object result)
            => dictionary.TryGetValue(binder.Name, out result);

        public override bool TryGetIndex(GetIndexBinder binder, object[] indexes, out object result)
        {
            var index = indexes[0];
            if (index is string) return dictionary.TryGetValue((string)index, out result);
            else if (index is int) result = dictionary.ElementAtOrDefault((int)index).Value;
            else result = null;
            return result != null;
        }

        public override IEnumerable<string> GetDynamicMemberNames()
            => this.dictionary.Keys;

    }
}
