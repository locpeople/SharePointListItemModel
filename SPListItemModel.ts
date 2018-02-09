import "reflect-metadata";
import {Web, sp} from "sp-pnp-js";
import moment from "moment-es6";

export interface ISPUrl {
    Description: string
    Url: string
}
export function SPList(name: string, site?: string): ClassDecorator {
    return target => {
        Reflect.defineMetadata("SPListName", name, target);
        if (site) Reflect.defineMetadata("SPListLocation", site, target)
    }
}

export function SPField(spfield?: string): PropertyDecorator {
    return (target, key) => {
        Reflect.defineMetadata(`SPField_${key}`, spfield ? spfield : key, target)
    }
}

function getSPFieldName(key: string, target) {
    return Reflect.getMetadata(key, target);
}

function getSPList(target) {
    const name = Reflect.getMetadata("SPListName", target);
    const web = Reflect.hasMetadata("SPListLocation", target)
        ? new Web(Reflect.getMetadata("SPListLocation", target))
        : sp.web;
    return web.lists.getByTitle(name).items;
}

function getMapper(target): { InternalName: string, ExternalName: string }[] {
    let output = [];
    let keys = Reflect.getMetadataKeys(target);
    keys = keys.filter(k => k.split("_")[0] == "SPField");
    keys.map(k => {
        output.push({
            InternalName: getSPFieldName(k, target),
            ExternalName: k.substr(8)
        })
    });
    return output;
}

export abstract class SPListItemModel {

    static getItemById<T extends SPListItemModel>(this: { new(): T }, id: number): Promise<T> {
        return getSPList(this).getById(id).get()
            .then(r => {
                let output = new this();
                output.rawData = r;
                output.thisType = this;
                return output;
            })
    }

    static getItemsByFilter<T extends SPListItemModel>(this: { new (): T }, query: string): Promise<T[]> {
        return getSPList(this).filter(query).get()
            .then(listdata => {
                let output = [];
                listdata.map(item => {
                    let thisitem = new this();
                    thisitem.rawData = item;
                    thisitem.thisType = this;
                    output.push(thisitem)
                })
                return output;
            })
    }

    static getAllItems<T extends SPListItemModel>(this: { new (): T }): Promise<T[]> {
        return getSPList(this).get()
            .then(listdata => {
                let output = [];
                listdata.map(item => {
                    let thisitem = new this();
                    thisitem.rawData = item;
                    thisitem.thisType = this;
                    output.push(thisitem)
                })
                return output;
            })
    }

    set thisType(type) {
        this._type = type;
    }

    private _type;

    set rawData(rawdata) {
        const mapper = getMapper(this);
        mapper.map(i => {
            const value = Reflect.getMetadata(`SPField_${i.ExternalName}`, this);
            let dataItem = rawdata[value];
            if (moment(dataItem, moment.ISO_8601, true).isValid()) dataItem = new Date(dataItem);
            this[i.ExternalName] = dataItem;
            this._cachedObj[i.InternalName] = dataItem;
        });
        if (this.ID == undefined) delete this.ID
    }

    private get _internalObj(): any {
        let output: any = {};
        getMapper(this).map(i => output[i.InternalName] = this[i.ExternalName]);
        if (output.ID == undefined) delete output.ID;
        return output;
    }

    private _cachedObj = {};

    private _hasChanged(postObj): Promise<boolean> {
        return SPListItemModel.getItemById.call(this._type, this.ID)
            .then(r => {
                let output = false;
                for (let item in postObj) {
                    let external = this._type.getExternalName(item);
                    if (this._cachedObj[item] !== r[external]) {
                        output = true;
                        break;
                    }
                }
                return output;
            })
    }

    static getInternalName<T extends SPListItemModel>(this: { new(): T }, ExternalFieldName: string) {
        const target = new this();
        return getSPFieldName(`SPField_${ExternalFieldName}`, target);
    }

    static getExternalName<T extends SPListItemModel>(this: { new(): T }, InternalFieldName: string) {
        const target = new this();
        let mapper = getMapper(target);
        let found = mapper.find(i=>i.InternalName == InternalFieldName);
        return found ? found.ExternalName : InternalFieldName;
    }

    submit(preventOverwrite = true): Promise<any> {
        let postobj = this._internalObj;
        for (let i in postobj) {
            if (!postobj[i]) delete postobj[i]
            else if (postobj[i] == this._cachedObj[i]) delete postobj[i]
        }
        delete postobj.ID;

        if (!this.ID) {
            return getSPList(this).add(postobj)
        }

        if (Object.keys(this._cachedObj).length === 0) {
            return getSPList(this).getById(this.ID).update(postobj)
        }

        let retprom = preventOverwrite
            ? this._hasChanged(postobj)
            : Promise.resolve(false);

        return retprom.then(c => {
            return c
                ? Promise.reject("Server data has changed, unable to update")
                : getSPList(this).getById(this.ID).update(postobj)
        })

    }

    @SPField()
    public ID: number;

}