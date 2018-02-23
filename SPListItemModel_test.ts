import "mocha";
import {assert} from "chai";
import moment from 'moment-es6';
import * as sinon from 'sinon';
import {SPListItemModel, SPField, SPList, ISPUrl} from "./SPListItemModel";
import * as pnp from 'sp-pnp-js';
import {SharePointQueryableShareableItem} from "sp-pnp-js/lib/sharepoint/sharepointqueryableshareable";

@SPList("TestList1", "http://testsite.com")
class TestModel1 extends SPListItemModel {
    @SPField()
    public TestField1: string;

    public TestField2: string;

    @SPField("InternalNameForTestField3")
    public TestField3: string;

    @SPField()
    public TestField4: Date;

    @SPField()
    public TestUrl: ISPUrl

}

@SPList("TestList2")
class TestModel2 extends SPListItemModel {

}

describe("Intiation and new data", () => {
    beforeEach(function () {
        this.add = sinon.stub(pnp.Items.prototype, "add")
            .resolves({ID: 1})
        this.tm1 = new TestModel1();
        this.tm2 = new TestModel2();
        this.tm1.TestField1 = "Test123";
        this.tm1.TestField2 = "Test456";
        this.tm1.TestField3 = "Test789";
    })
    afterEach(function () {
        this.add.restore();
    })


    it("Models initialize without errors", function () {
        assert.ok(this.tm1, "Unable to initalize model class (1)");
        assert.ok(this.tm2, "Unable to initalize model class (2)")
    })

    it("Calls pnp's add method correctly", function (done) {
        let add = this.add;


        this.tm1.submit()
            .then(() => {
                let postobj = {
                    TestField1: "Test123",
                    InternalNameForTestField3: "Test789"
                };

                assert(add.calledOnce, "Submit method not called");
                assert(add.calledWith(postobj), "Incorrect data supplied to submit method when adding new record");
                done();
            })
            .catch(e => done(new Error(e)))
    });

    it("Correctly returns the internal field name", function () {
        assert.equal(TestModel1.getInternalName("TestField3"), "InternalNameForTestField3")
    })


});

describe("Working with existing data", () => {
    beforeEach(function () {
        this.resolveobj = {
            TestField1: "Test123",
            InternalNameForTestField3: "Test321",
            IrrelevantField: "Test567",
            TestField4: "2018-01-25T00:00:00Z",
            TestUrl: {Description: "this is an urlnode", Url: "http://www.test.com"},
            ID: 1
        };
        this.getbyid = sinon.stub(pnp.ODataQueryable.prototype, "get")
            .resolves(this.resolveobj);
        this.update = sinon.stub(pnp.Item.prototype, "update")
            .resolves({})
    });
    afterEach(function () {
        this.getbyid.restore();
        this.update.restore();
    });
    it("Creates a correct new object from a SharePoint response", function (done) {
        TestModel1.getItemById(1)
            .then((r: TestModel1) => {
                assert.ok(r, "Model object not loaded");
                assert.equal(r.TestField1, "Test123", "TestField1 not populated correctly");
                assert.equal(r.TestField3, "Test321", "TestField2 not populated correctly");
                assert.typeOf(r.TestField4, "Date", "TestField4 not correctly parsed to Date object");
                assert(moment(r.TestField4).isSame(new Date("2018-01-25T00:00:00Z")), "TestField4 has incorrect date");
                assert.hasAllKeys(r.TestUrl, ["Description", "Url"])
                done();
            })
            .catch(done);
    })


    it("Updates correctly", function (done) {
        TestModel1.getItemById(1)
            .then(obj => {
                obj.TestField1 = "Changed";
                obj.TestField3 = "Also changed";
                return obj.submit()
            })
            .then(() => {
                assert(this.update.calledWith({
                    TestField1: "Changed",
                    InternalNameForTestField3: "Also changed"
                }), "Incorrect data sent to submit while updating record");
                done();
            })
            .catch(e => done(new Error(e)))
    });


    it("Refuses to update when server data has changed", function (done) {

        let promise = TestModel1.getItemById(1)
            .then(obj => {
                    obj.TestField1 = "Changed";
                    this.resolveobj.TestField1 = "ChangedAgain";
                    return obj.submit();
                }
            );
        promise
            .then(
                () => done(new Error("Submit proceeded when server data had changed")),
                e => {
                    assert.equal(e, "Server data has changed, unable to update", "Update is refused for the wrong reason");
                    done();
                }
            )
    })

    it("Forces overwrite with preventOverwrite:false", function (done) {
        let promise = TestModel1.getItemById(1)
            .then(obj => {
                    obj.TestField1 = "Changed";
                    this.resolveobj.TestField1 = "ChangedAgain";
                    return obj.submit(false);
                }
            );
        promise.then(r => {
            assert(true, "Update not correctly done with preventOverwrite:false");
            done();
        })
    })

    it("Updates a listitem that hasn't been loaded", function (done) {
        let item = new TestModel1();
        item.TestField1 = "Test123";
        item.ID = 1;
        let update = this.update;
        item.submit()
            .then(r => {
                assert.ok(r, "Unexpected server resposne when updating unloaded listitem");
                let a = update.calledWith({
                    TestField1: "Test123"
                });
                assert(a, "Update on unloaded item unsuccessful");
                done();
            })
            .catch(e => {
                done(new Error(e));
            })
    })
});

describe("Deleting data", () => {
    beforeEach(function () {
        //SharePointQueryableShareableItem
        this.delete = sinon.stub(pnp.Item.prototype, "delete")
            .resolves()
    });
    it("Correctly deletes by id", function (done) {
        TestModel1.deleteItemById(1)
            .then(() => {
                assert(this.delete.calledOnce, "Delete method not called correctly");
                done();
            })
            .catch(e => done(new Error(e)))
    })
});

describe("Lists of data", () => {
    beforeEach(function () {
        this.resolveobj1 = {
            TestField1: "Test123",
            InternalNameForTestField3: "Test321",
            TestField4: "2018-01-25T00:00:00Z",
            ID: 1
        };
        this.resolveobj2 = {
            TestField1: "Test123",
            InternalNameForTestField3: "Test567890",
            TestField4: "2018-01-26T00:00:00Z",
            ID: 2
        };
        this.filter = sinon.stub(pnp.ODataQueryable.prototype, "get")
            .resolves([this.resolveobj1, this.resolveobj2]);
    });
    afterEach(function () {
        this.filter.restore();
    });
    it("Correctly requests and parses all records in the list", function (done) {
        TestModel1.getAllItems()
            .then(r => {
                assert.ok(r, "No sane response")
                assert(r.length == 2, "Returned array has the wrong length")
                assert(r[0].TestField3 == "Test321", "First item in the array is not correct");
                assert(r[1].TestField3 == "Test567890", "Second item in the array is not correct");
                done();
            })
            .catch(e => {
                done(new Error(e));
            })
    })
    it("Correctly requests and parses data with a filter", function (done) {
        TestModel1.getItemsByFilter("Testfield1 eq Test123")
            .then(r => {
                assert.ok(r, "No sane response")
                assert(r.length == 2, "Returned array has the wrong length")
                assert(r[0].TestField3 == "Test321", "First item in the array is not correct");
                assert(r[1].TestField3 == "Test567890", "Second item in the array is not correct");
                done();
            })
            .catch(e => {
                done(new Error(e));
            })
    })
});