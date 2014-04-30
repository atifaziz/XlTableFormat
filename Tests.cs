namespace Tester
{
    using System.Reflection;
    using System.Runtime.InteropServices;
    using NUnit.Framework;

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
    }
}
