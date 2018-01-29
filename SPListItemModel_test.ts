import {assert} from 'chai';
import * as moment from 'moment';
import * as sinon from 'sinon';
import {SPListItemModel, SPField, SPList} from "./SPListItemModel";
import * as pnp from 'sp-pnp-js';

@SPList("TestList1", "http://testsite.com")
class TestModel1 extends SPListItemModel {
    @SPField()
    public TestField1: string;

    public TestField2: string;

    @SPField("InternalNameForTestField3")
    public TestField3: string;

    @SPField()
    public TestField4: Date;
}

@SPList("TestList2")
class TestModel2 extends SPListItemModel {

}

describe("Intiation and new data", () => {
    beforeEach(() => {
        this.add = sinon.stub(pnp.Items.prototype, "add")
            .resolves({ID: 1})
        this.tm1 = new TestModel1();
        this.tm2 = new TestModel2();
        this.tm1.TestField1 = "Test123";
        this.tm1.TestField2 = "Test456";
        this.tm1.TestField3 = "Test789";
    })
    afterEach(() => {
        this.add.restore();
    })


    it("Models initialize without errors", () => {
        assert.ok(this.tm1, "Unable to initalize model class (1)");
        assert.ok(this.tm2, "Unable to initalize model class (2)")
    })

    it("Calls pnp's add method correctly", done => {
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

});

describe("Working with existing data", () => {
    beforeEach(() => {
        this.resolveobj = {
            TestField1: "Test123",
            InternalNameForTestField3: "Test321",
            IrrelevantField: "Test567",
            TestField4: "2018-01-25T00:00:00Z",
            ID: 1
        };
        this.getbyid = sinon.stub(pnp.ODataQueryable.prototype, "get")
            .resolves(this.resolveobj);
        this.update = sinon.stub(pnp.Item.prototype, "update")
            .resolves({})
    });
    afterEach(() => {
        this.getbyid.restore();
        this.update.restore();
    });
    it("Creates a correct new object from a SharePoint response", done => {
        TestModel1.getItemById(1)
            .then((r: TestModel1) => {
                assert.ok(r, "Model object not loaded");
                assert.equal(r.TestField1, "Test123", "TestField1 not populated correctly");
                assert.equal(r.TestField3, "Test321", "TestField2 not populated correctly");
                assert.typeOf(r.TestField4, "Date", "TestField4 not correctly parsed to Date object");
                assert(moment(r.TestField4).isSame(new Date("2018-01-25T00:00:00Z")), "TestField4 has incorrect date");
                done();
            })
            .catch(done);
    })


    it("Updates correctly", done => {
        TestModel1.getItemById(1)
            .then(obj => {
                obj.TestField1 = "Changed";
                return obj.submit()
            })
            .then(() => {
                assert(this.update.calledWith({
                    TestField1: "Changed",
                }), "Incorrect data sent to submit while updating record");
                done();
            })
            .catch(e => done(new Error(e)))
    });


    it("Refuses to update when server data has changed", done => {

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

    it("Forces overwrite with preventOverwrite:false", done => {
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

    it("Updated a listitem that hasn't been loaded", done => {
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