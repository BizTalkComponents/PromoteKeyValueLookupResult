using BizTalkComponents.Utilities.LookupUtility.Repository;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using Winterdom.BizTalk.PipelineTesting;

namespace BizTalkComponents.PipelineComponents.PromoteKeyValueLookupResult.Tests.UnitTests
{
    [TestClass]
    public class PromoteKeyValueLookupResultTests
    {
        private readonly Mock<ILookupRepository> mock = new Mock<ILookupRepository>();
        private readonly Dictionary<string, string> d = new Dictionary<string, string>();
        

        [TestInitialize]
        public void Init()
        {
            d.Add("TestKey", "TestValue");

            mock.Setup(l => l.LoadList("TestList")).Returns(d);
        }

        [TestMethod]
        public void TestPromoteKeyValueLookupResult()
        {
            var component = new SharepointPromoteKeyValueLookupResult(mock.Object)
            {
                DestinationPropertyPath = "NS#DestProp",
                ListName = "TestList",
                SourcePropertyPath = "NS#SrcProp"
            };

            var pipeline = PipelineFactory.CreateEmptyReceivePipeline();
            string m = "<body></body>";

            var message = MessageHelper.CreateFromString(m);

            message.Context.Promote("SrcProp", "NS", "TestKey");

            pipeline.AddComponent(component, PipelineStage.Decode);

            var result = pipeline.Execute(message);


            Assert.AreEqual("TestValue", result[0].Context.Read("DestProp", "NS").ToString());   
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void TestPromoteKeyValueLookupResultInvalidList()
        {
            var component = new SharepointPromoteKeyValueLookupResult(mock.Object)
            {
                DestinationPropertyPath = "NS#DestProp",
                ListName = "NonExisting",
                SourcePropertyPath = "NS#SrcProp"
            };

            var pipeline = PipelineFactory.CreateEmptyReceivePipeline();
            string m = "<body></body>";

            var message = MessageHelper.CreateFromString(m);

            message.Context.Promote("SrcProp", "NS", "TestKey");

            pipeline.AddComponent(component, PipelineStage.Decode);

            var result = pipeline.Execute(message);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void TestPromoteKeyValueLookupResultInvalidKey()
        {
            var component = new SharepointPromoteKeyValueLookupResult(mock.Object)
            {
                DestinationPropertyPath = "NS#DestProp",
                ListName = "TestList",
                SourcePropertyPath = "NS#SrcProp"
            };

            var pipeline = PipelineFactory.CreateEmptyReceivePipeline();
            string m = "<body></body>";

            var message = MessageHelper.CreateFromString(m);

            message.Context.Promote("SrcProp", "NS", "NonExisting");

            pipeline.AddComponent(component, PipelineStage.Decode);

            var result = pipeline.Execute(message);
        }

    }
}
