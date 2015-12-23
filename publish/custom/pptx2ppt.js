// =============================================================================
self.pptx2ppt = {
	// ————————————————————————————————————————————————————————————————————————————————
	Intelledox : {
		init : function(){
			//console.info('pptx2ppt.Intelledox.init');
			// DocumentDownload.ashx/PowerPoint%20Example.pptx?FileId=9ae15a30-72e8-4a63-aca8-2ab631a6c0f1&Extension=.pptx&JobId=38670273-5908-4508-b591-9cab1f6a4d95
			$('a.fileDownloadLink[href^="DocumentDownload.ashx"]').each( pptx2ppt.Intelledox.makeLink );
		},
		makeLink : function(idx,el){
			//console.info('pptx2ppt.Intelledox.makeLink');
			//console.log('idx',idx);
			//console.log('el',el);
            /*
		    var href = '/pptx2ppt/Default?source=' + escape(el.href);
            */
		    var arrVars = {};
		    try {
		        var arrPairs = el.href.split('?')[1].split('&');
		        for (i in arrPairs) {
		            if (arrPairs[i].search(/\=/) > -1) {
		                arrPair = arrPairs[i].split("=");
		                arrVars[arrPair[0]] = arrPair[1];
		            }
		        }
		    } catch (err) {
		        console.error('err',err);
		    }
		    console.log('arrVars', arrVars);
		    console.log("arrVars['FileId']=", arrVars['FileId']);
		    console.log("arrVars['JobId']=", arrVars['JobId']);

		    var href = '/pptx2ppt/Default?' + escape( 'FileId=' + arrVars['FileId'] + '&JobId=' + arrVars['JobId']);

            // Inject customization
		    $(el).after(
				' <a href="' + href + '">(ppt)</a>'
			);
		}
	}
	// ————————————————————————————————————————————————————————————————————————————————
};
// =============================================================================
// boot
$(document).ready(  pptx2ppt.Intelledox.init  );
// =============================================================================