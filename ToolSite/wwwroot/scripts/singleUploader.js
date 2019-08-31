define(["dojo/_base/declare", "dijit/_WidgetBase", "dijit/_TemplatedMixin", "dojo/text!./templates/singleUploader.html", "dojo/on","dojo/dom-attr","dojo/dom-class"], function (declare, _WidgetBase, _TemplatedMixin, template, on,domAttr,domClass) {

    return declare([_WidgetBase, _TemplatedMixin], {
        templateString: template,
        value:'',
        validate: function(){
            if(this.required&&this.value==='') return false;
            return true;
        },
        focus:function(){
            domAttr.set(this.displayTextNode,'placeholder',"请选取上传的文件");
            domClass.add(this.displayTextNode,'required');
        },
        postCreate: function () {
            var _this=this;
            on(_this.displayTextNode,'click',function(){_this.hiddenFileNode.click();});
            on(_this.hiddenFileNode,'change',function(){
                domClass.remove(_this.displayTextNode,'required');
               var fileNameArr= _this.hiddenFileNode.value.split('\\');
               _this.value=fileNameArr[fileNameArr.length-1];
               _this.displayTextNode.value=_this.value;
               domAttr.set(_this.displayTextNode,'title',_this.hiddenFileNode.value);  
            });
        }
    });//return
});//define