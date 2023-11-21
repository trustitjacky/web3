//selaa

function selaa(x,y){

		var ajaxClickA=new Ajax.Updater('cc2','cc2.asp',{method:'post',parameters:'fid='+x+'&selid='+y,evalScripts:true});

}

function chg_admin_dep(x){
		var ajaxClickA=new Ajax.Updater('chg_adm_dep','chg_adm_dep.asp',{method:'post',parameters:'x='+x,evalScripts:true});
}

