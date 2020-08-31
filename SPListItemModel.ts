import "reflect-metadata";
import { Web, sp } from "@pnp/sp";
import moment from 'moment-es6';

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
        Reflect.defineMetadata(`SPField_${<string>key}`, spfield ? spfield : key, target)
    }
}

function getSPFieldName(key: string, target) {
    return Reflect.getMetadata(key, target);
}

function isSPField(key: string, target) {
    return Reflect.hasMetadata(`SPField_${key}`, target);
}


function getSPList(target) {
    const name = Reflect.getMetadata("SPListName", target);
    const web = Reflect.hasMetadata("SPListLocation", target)
        ? new Web(Reflect.getMetadata("SPListLocation", target))
        : sp.web;
    return web.lists.getByTitle(name).items;
}

function getMapper(target): { InternalName: string, ExternalName: string }[] {
    let keys = Reflect.getMetadataKeys(target);
    keys = keys.filter(k => k.split("_")[0] == "SPField");
    return keys.map(k => {
        return {
            InternalName: getSPFieldName(k, target),
            ExternalName: k.substr(8)
        }
    });
}

export abstract class SPListItemModel {


    static getItemById<T extends SPListItemModel>(this: { new(): T }, id: number): Promise<T> {
        return getSPList(this).getById(id).get()
            .then(r => {
                let output = new this();
                output.rawData = r;
                return output;
            })
    }


    static getItemsByFilter<T extends SPListItemModel>(this: { new(): T }, query: string, top: number = null): Promise<T[]> {

        return getSPList(this).filter(query).top(top ? top : 500).get()
            .then(listdata => {
                return listdata.map(item => {
                    let thisitem = new this();
                    thisitem.rawData = item;
                    return thisitem;
                })
            })
    }



    static deleteItemById<T extends SPListItemModel>(this: { new(): T }, id: number): Promise<void> {
        return getSPList(this).getById(id).delete()
    }


    static getAllItems<T extends SPListItemModel>(this: { new(): T }): Promise<T[]> {
        return getSPList(this).getAll()
            .then(listdata => {
                return listdata.map(item => {
                    let thisitem = new this();
                    thisitem.rawData = item;
                    return thisitem;
                })
            })
    }

    static getInternalName<T extends SPListItemModel>(this: { new(): T }, ExternalFieldName: string) {
        const target = new this();
        return getSPFieldName(`SPField_${ExternalFieldName}`, target);
    }


    set rawData(rawdata) {
        const mapper = getMapper(this);
        mapper.forEach(i => {
            const value = Reflect.getMetadata(`SPField_${i.ExternalName}`, this);
            let dataItem = rawdata[value];
            if (parseInt(dataItem).toString() != dataItem && moment(dataItem, moment.ISO_8601, true).isValid()) dataItem = new Date(dataItem);
            this[i.ExternalName] = dataItem;
            this._cachedObj[i.InternalName] = dataItem;
        });
        if (this.ID == undefined) delete this.ID
    }

    private get _internalObj(): any {
        let output: any = {};
        getMapper(this).forEach(i => output[i.InternalName] = this[i.ExternalName]);
        if (output.ID == undefined) delete output.ID;
        return output;
    }

    private _cachedObj = {};

    private _hasChanged(postObj): Promise<boolean> {
        return SPListItemModel.getItemById.call(this.constructor, this.ID)
            .then(r => {
                let output = false;
                let mapper = getMapper(this.constructor.prototype);

                for (let item in mapper) {
                    let thisitem = mapper[item];
                    let cacheditem = this._cachedObj[thisitem.InternalName];
                    let newitem = r[thisitem.ExternalName];

                    let same = cacheditem == newitem;

                    if (!same && cacheditem instanceof Date) {
                        same = moment(cacheditem).isSame(newitem);
                    }

                    if (!same && typeof cacheditem == "object") {
                        same = true;
                        for (let i in cacheditem) {
                            if (cacheditem[i] !== newitem[i]) {
                                same = false;
                                break;
                            }
                        }
                    }

                    if (!same) {
                        output = true;
                        break;
                    }

                }

                return output;
            })
    }

    submit(preventOverwrite = true): Promise<any> {
        let postobj = this._internalObj;
        let list = getSPList(this.constructor);
        for (let i in postobj) {
            if (postobj[i] == undefined) delete postobj[i]
            else if (postobj[i] == this._cachedObj[i]) delete postobj[i]
        }
        delete postobj.ID;

        if (!this.ID) {
            return list.add(postobj)
        }

        if (Object.keys(this._cachedObj).length === 0) {
            return list.getById(this.ID).update(postobj)
        }

        let retprom = preventOverwrite
            ? this._hasChanged(postobj)
            : Promise.resolve(false);

        return retprom.then(c => {
            return c
                ? Promise.reject("Server data has changed, unable to update")
                : list.getById(this.ID).update(postobj)
        })

    }

    @SPField()
    public ID: number;

}
