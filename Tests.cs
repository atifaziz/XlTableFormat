#region Copyright (c) 2014 Atif Aziz. All rights reserved.
//
// Copyright (c) 2014 Atif Aziz. All rights reserved.
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
#endregion

namespace Tester
{
    #region Imports

    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using NUnit.Framework;

    #endregion

    [TestFixture]
    sealed class Tests
    {
        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Initialized()
        {
            Assert.That(XlTableFormat.DefaultDataFactory, Is.Not.Null);
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Blank()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Blank, Is.Null);
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Skip()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Skip, Is.EqualTo(Missing.Value));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Table()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Table(12, 34), Is.EqualTo(new[] { 12, 34 }));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Float()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Float(12.34), Is.EqualTo(12.34));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_String()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.String("foobar"), Is.EqualTo("foobar"));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_BoolTrue()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Bool(true), Is.True);
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_BoolFalse()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Bool(false), Is.False);
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Error()
        {
            var result = XlTableFormat.DefaultDataFactory.Error(42);
            Assert.That(result, Is.Not.Null);
            Assert.That(result, Is.InstanceOf<ErrorWrapper>());
            Assert.That(((ErrorWrapper) result).ErrorCode, Is.EqualTo(42));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void DefaultDataFactory_Int()
        {
            Assert.That(XlTableFormat.DefaultDataFactory.Int(42), Is.EqualTo(42));
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void Read_Strings()
        {
            using (var r = Read(
                0x10, 0x00, 0x04, 0x00, 0x01, 0x00, 0x03, 0x00, 
                0x02, 0x00, 0x0c, 0x00, 0x03, 0x66, 0x6f, 0x6f, 
                0x03, 0x62, 0x61, 0x72, 0x03, 0x62, 0x61, 0x7a))
            {
                Assert.That(r.Read(), Is.EqualTo(new[] { 1, 3 }));
                Assert.That(r.Read(), Is.EqualTo("foo"));
                Assert.That(r.Read(), Is.EqualTo("bar"));
                Assert.That(r.Read(), Is.EqualTo("baz"));
            }         
        }

        static Reader<object> Read(params byte[] bytes)
        {
            return Reader.Create(XlTableFormat.Read(new MemoryStream(bytes, writable: false)));
        }

        static class Reader
        {
            public static Reader<T> Create<T>(IEnumerator<T> enumerator)
            {
                return new Reader<T>(enumerator);
            }
        }

        sealed class Reader<T> : IDisposable
        {
            IEnumerator<T> _enumerator;

            public Reader(IEnumerator<T> enumerator)
            {
                if (enumerator == null) throw new ArgumentNullException("enumerator");
                _enumerator = enumerator;
            }

            public T Read()
            {
                var e = _enumerator;
                if (e == null)
                    throw new ObjectDisposedException("Reader");
                if (!e.MoveNext())
                    throw new InvalidOperationException();
                return e.Current;
            }

            public void Dispose()
            {
                var e = _enumerator;
                if (e == null) 
                    return;
                _enumerator = null;
                var eof = !e.MoveNext();
                e.Dispose();
                if (eof) return;
                // Deliberate violation of the rule that Dispose should now throw
                throw new Exception("Source was not fully read.");
            }
        }
    }
}
