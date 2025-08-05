using System;
using System.Runtime.Serialization;

namespace OfficeIMO.Converters {
    [Serializable]
    public class ConversionException : Exception {
        public ConversionException() {
        }

        public ConversionException(string message) : base(message) {
        }

        public ConversionException(string message, Exception innerException) : base(message, innerException) {
        }

        protected ConversionException(SerializationInfo info, StreamingContext context) : base(info, context) {
        }
    }
}

