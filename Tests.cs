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
    using NUnit.Framework.Constraints;

    #endregion

    // ReSharper disable AccessToStaticMemberViaDerivedType

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
                Assert.That(r.EOF, Is.True);
            }
        }

        [Test] // ReSharper disable once InconsistentNaming
        public void Read_Big()
        {
            // 01/05/2014 | -0.04 | -0.08 | TRUE  | -0.08
            //       0.45 | -0.23 | -0.38 | -0.36 | -0.24
            //       0.12 |  0.06 |  0.48 |       | 
            //            |       |  0.38 |  0.11 |  0.25
            //       0.00 |  0.23 | -0.03 | -0.05 | -0.08
            //       0.30 | -0.19 |  0.08 |  0.23 | -0.10
            //      -0.17 |  0.23 | #N/A  |  0.11 | -0.29
            //       0.26 | -0.48 |  0.34 | -0.02 |  0.44
            //       0.02 |  0.43 |  0.42 |  0.10 | -0.14
            //       0.13 |  0.46 |  0.33 |  0.49 |  0.19
            //      -0.05 |  0.35 |  0.40 |  0.05 |  0.36
            //       0.17 |  0.13 |  0.44 |  0.04 | -0.27
            //      -0.17 |  0.36 |  0.25 |  0.24 | MAY

            using (var r = Read(
                0x10, 0x00, 0x04, 0x00, 0x0d, 0x00, 0x05, 0x00, 0x01, 0x00, 0x18, 0x00, 0x07, 0x90, 0x3e, 0x1c,
                0x0d, 0x64, 0xe4, 0x40, 0xe0, 0x1c, 0x21, 0x2e, 0x17, 0xb4, 0xa4, 0xbf, 0x18, 0x64, 0xb1, 0x0e,
                0x27, 0x87, 0xb4, 0xbf, 0x03, 0x00, 0x02, 0x00, 0x01, 0x00, 0x01, 0x00, 0x48, 0x00, 0x18, 0x39,
                0x39, 0x7a, 0x04, 0x2e, 0xb5, 0xbf, 0xc0, 0x44, 0x05, 0xdd, 0x10, 0x7d, 0xdc, 0x3f, 0x50, 0xa5,
                0xbe, 0x87, 0x0f, 0xea, 0xcc, 0xbf, 0xf0, 0x03, 0xf4, 0x2c, 0x26, 0x73, 0xd8, 0xbf, 0x24, 0x8f,
                0x2d, 0x01, 0x53, 0x0b, 0xd7, 0xbf, 0x7c, 0x41, 0x1c, 0x44, 0x92, 0xe3, 0xce, 0xbf, 0x30, 0x7e,
                0x19, 0x45, 0xd0, 0xf2, 0xbf, 0x3f, 0xb0, 0xb3, 0xce, 0xd8, 0x4d, 0x3c, 0xac, 0x3f, 0x18, 0x33,
                0x0b, 0x38, 0x91, 0x93, 0xde, 0x3f, 0x05, 0x00, 0x02, 0x00, 0x04, 0x00, 0x01, 0x00, 0x78, 0x00,
                0xbe, 0x22, 0xaf, 0xa8, 0xc7, 0x61, 0xd8, 0x3f, 0xe0, 0x8b, 0x65, 0xd7, 0xb6, 0xf1, 0xbb, 0x3f,
                0x82, 0xd2, 0xdc, 0xec, 0xfe, 0x44, 0xd0, 0x3f, 0x00, 0xd4, 0xcc, 0x84, 0x16, 0xc5, 0x48, 0x3f,
                0x64, 0xed, 0x80, 0x6e, 0x1a, 0x7c, 0xcd, 0x3f, 0x00, 0xad, 0xd0, 0x54, 0x86, 0x42, 0x9e, 0xbf,
                0x00, 0xf0, 0x44, 0xf7, 0x8b, 0xe0, 0xaa, 0xbf, 0x18, 0x12, 0xed, 0x35, 0x0b, 0xad, 0xb5, 0xbf,
                0x6e, 0x0e, 0xe5, 0x21, 0x4e, 0x49, 0xd3, 0x3f, 0xf4, 0xd1, 0x48, 0x55, 0x2a, 0xdf, 0xc7, 0xbf,
                0x70, 0x1a, 0xb6, 0x0a, 0x1e, 0x39, 0xb4, 0x3f, 0x9c, 0x80, 0x54, 0x6e, 0x2b, 0x16, 0xcd, 0x3f,
                0xd0, 0xd2, 0xfb, 0x31, 0x1d, 0xab, 0xba, 0xbf, 0x44, 0xc8, 0xd6, 0xb2, 0x3a, 0x3b, 0xc6, 0xbf,
                0xe4, 0x77, 0x9a, 0xb7, 0xa3, 0xeb, 0xcc, 0x3f, 0x04, 0x00, 0x02, 0x00, 0x2a, 0x00, 0x01, 0x00,
                0xf8, 0x00, 0x20, 0x75, 0x6b, 0x82, 0x82, 0x53, 0xbc, 0x3f, 0x62, 0x5e, 0xca, 0x78, 0xd0, 0xd9,
                0xd2, 0xbf, 0xc8, 0x6a, 0xa4, 0x89, 0xe8, 0x65, 0xd0, 0x3f, 0x10, 0x04, 0xf1, 0x7a, 0x6f, 0x02,
                0xdf, 0xbf, 0xea, 0x98, 0xc4, 0xec, 0x18, 0xac, 0xd5, 0x3f, 0x80, 0x3f, 0xe4, 0xd8, 0xeb, 0xfd,
                0x8f, 0xbf, 0xc8, 0xe2, 0x6e, 0xbd, 0xa3, 0x14, 0xdc, 0x3f, 0xc0, 0x38, 0x1a, 0xaf, 0xc8, 0x11,
                0x97, 0x3f, 0x00, 0xaa, 0x1e, 0xf3, 0xd7, 0xcf, 0xdb, 0x3f, 0x6e, 0x5a, 0xdf, 0x89, 0x6c, 0xfb,
                0xda, 0x3f, 0x68, 0x4e, 0x0c, 0x3a, 0x81, 0x58, 0xb8, 0x3f, 0x28, 0xcc, 0x70, 0x59, 0x1c, 0xc9,
                0xc1, 0xbf, 0x38, 0x78, 0x96, 0xea, 0xeb, 0x2d, 0xc0, 0x3f, 0x44, 0x33, 0x03, 0xc7, 0x38, 0x2d,
                0xdd, 0x3f, 0x66, 0x6c, 0x8c, 0xaa, 0x5c, 0x0d, 0xd5, 0x3f, 0x0c, 0xb4, 0x87, 0x15, 0x91, 0x4a,
                0xdf, 0x3f, 0x54, 0x6c, 0xab, 0x94, 0x7f, 0x18, 0xc8, 0x3f, 0x60, 0x02, 0x72, 0x07, 0x7a, 0xce,
                0xaa, 0xbf, 0x08, 0xb7, 0x51, 0xcd, 0x7f, 0x69, 0xd6, 0x3f, 0x2c, 0xf9, 0x25, 0x5b, 0x6e, 0x61,
                0xd9, 0x3f, 0xe0, 0x33, 0x16, 0x73, 0x0c, 0x31, 0xa7, 0x3f, 0x50, 0xcf, 0xd1, 0xf3, 0xe2, 0x03,
                0xd7, 0x3f, 0xe8, 0x5e, 0x4e, 0x23, 0x53, 0x49, 0xc6, 0x3f, 0xa4, 0x18, 0xdb, 0x3b, 0x9a, 0x49,
                0xc0, 0x3f, 0x14, 0xb3, 0x08, 0xfa, 0x1b, 0x66, 0xdc, 0x3f, 0x30, 0xcd, 0x05, 0x6b, 0xf4, 0xa4,
                0xa2, 0x3f, 0x5c, 0xc3, 0xdb, 0xfd, 0xc6, 0x0b, 0xd1, 0xbf, 0xf0, 0x86, 0x21, 0xb9, 0x4d, 0x5e,
                0xc6, 0xbf, 0x2e, 0x6c, 0xd8, 0xdf, 0xe2, 0x4a, 0xd7, 0x3f, 0x7a, 0x0f, 0xe3, 0x76, 0x33, 0x2d,
                0xd0, 0x3f, 0x54, 0x40, 0x43, 0x5b, 0x25, 0x1a, 0xce, 0x3f, 0x02, 0x00, 0x04, 0x00, 0x03, 0x4d,
                0x41, 0x59))
            {
                Assert.That(r.Read(), Is.EqualTo(new[] { 13, 5 }));
                
                Assert.That(r.Read(), Is.RoughlyEqualTo(41760, 0));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.04, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.08, 2));
                Assert.That(r.Read(), Is.True);
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.08, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.45, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.23, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.38, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.36, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.24, 2));
                
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.12, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.06, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.48, 2));
                Assert.That(r.Read(), Is.Null);
                Assert.That(r.Read(), Is.Null);

                Assert.That(r.Read(), Is.Null);
                Assert.That(r.Read(), Is.Null);
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.38, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.11, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.25, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.00, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.23, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.03, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.05, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.08, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.30, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.19, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.08, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.23, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.10, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.17, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.23, 2));
                Assert.That(((ErrorWrapper) r.Read()).ErrorCode, Is.EqualTo(42));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.11, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.29, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.26, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.48, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.34, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.02, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.44, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.02, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.43, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.42, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.10, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.14, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.13, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.46, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.33, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.49, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.19, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.05, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.35, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.40, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.05, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.36, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.17, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.13, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.44, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.04, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.27, 2));

                Assert.That(r.Read(), Is.RoughlyEqualTo(-0.17, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.36, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.25, 2));
                Assert.That(r.Read(), Is.RoughlyEqualTo( 0.24, 2));
                Assert.That(r.Read(), Is.EqualTo("MAY"));

                Assert.That(r.EOF, Is.True);
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

            // ReSharper disable once InconsistentNaming
            public bool EOF { get { return !_enumerator.MoveNext(); } }

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
                e.Dispose();
            }
        }
    }

    sealed class Is : NUnit.Framework.Is
    {
        public static RoundingConstraint RoughlyEqualTo(double expected, int decimals)
        {
            return new RoundingConstraint(expected, decimals);
        }
    }

    sealed class RoundingConstraint : Constraint
    {
        readonly double _expected;
        readonly int _decimals;

        public RoundingConstraint(double expected, int decimals)
        {
            _expected = expected;
            _decimals = decimals;
        }

        // ReSharper disable once ParameterHidesMember
        public override bool Matches(object actual)
        {
            base.actual = actual;
            return actual is double 
                && Math.Abs(_expected - Math.Round((double) actual, _decimals)) < double.Epsilon;
        }

        public override void WriteDescriptionTo(MessageWriter writer)
        {
            writer.WritePredicate(string.Format(
                @"approximately equal to (rounded to {0} decimal{1})", 
                _decimals, 
                _decimals == 1 ? null : "s"));
            writer.WriteExpectedValue(_expected);
        }
    }
}
