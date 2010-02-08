/* 

Script: Spreadsheet.js
	Spreadsheet -  Wave Spreadsheet Gadget that enables spreadsheet integration in your wave

Copyright:
	Copyright (c) 2009 David Buchbut, <http://spreadsheet.happinesbeats.com/>.

License:
	MIT-style license.

Contributors:

    - David Buchbut

*/



var Spreadsheet = new Class({
    initialize: function(){
        this.activeElement = {row:1,cell:1};
        this.gridSize = {rows:8,cells:26,frozenRows:0};
        this.markerColor = ['green','brown','orange','red','purple','lightblue','yellow','pink']
        this.advancedFormula = 'SUM,AVERAGE,PRODUCT,MIN,MAX';
        this.formulaHandler = new AdvancedFormula(this);
        this.resizeMargine = 10; // the margine to start show resize
        this.colorPalatte = [['000000','2C2C2C','838383','969696','c7C7C7','E1e1e1','FFFFFF'],
				['FCB218','FFE300','008A45','5DF9F9','7395FC','CD54FC','FB0F0F'],
				['FFCC00','FFFF00','00FF00','00FFFF','00ccff','993366','FF00FF'],
				['FADCB3','FFFF99','CCFFCC','CCFFFF','C2D1f0','E1c7e1','FF99CC'],
				['FFCC99','EBD780','B3D580','BDE6E1','99CCff','CC99FF','E69999'],
				['FF6600','808000','008000','008080','0000FF','6666CC','FF0000'],
				['993300','333300','003300','003366','000080','333399','800000']]
        this.localParticipantColor = 'blue';
        this.viewerId;
        this.participants;
        this.actionIndex = 0;
        this.sheet = 0;
        this.divHeaderContainer;
        this.actions = {};
        this.currentRange = null; // selected range
        this.grids = $('divContainer');
        this.matrix = this.grids.getElement('.divMatrix');
        this.divHeaderContainer = this.grids.getElement('div.divHeaderContainer');
        this.frozRowHead = this.grids.getElement('div.divFrozenPanel .rowMarker');
        this.rowHead = this.grids.getElement('table.framesContainer .rowMarker');
        this.cellLocationTd = $('cellLocationTd');
        this.head = {
                'grid':this.grids.getElement('.headGrid'),
                'gridContainer':$('divHeaderContainer'),
                'currentCell':$('divHeaderContainer').getElement('.divCurrentCell'),
                'rowHeader':this.grids.getElement('div.divFrozenPanel .rowMarker')
        };   // pointer to elements in the header

        this.main = {
                'grid':this.grids.getElement('.tblHeaderContent'),
                'gridContainer':this.grids.getElement('.divMainArea'),
                'currentCell':this.grids.getElement('.divMainArea').getElement('.divCurrentCell'),
                'rowHeader':this.grids.getElement('table.framesContainer .rowMarker'),
                'rowHeaderContainer': $('divHeaderRow')
        };   // pointer to elements in the header
        this.ghostFreezebar = this.grids.getElement('.ghostFreezeBar');
        this.columnResizer = this.positionColumnResizer();
        this.selObj = {}   // pointer to selection objects
        this.selObj.selection = this.matrix.getElement('.divRange');
        this.selObj.formulaSelector = this.matrix.getElement('.formulaMarker');
        this.editBox = this.grids.getElement('.editBox');
        this.editBoxContainer = this.grids.getElement('.editBoxContainer');
        this.head.mouseOvercell = {row:null, cell:null};
        this.head.mouseDown = false;
        this.renderFrozenTable();
        this.renderMainTable();
        this.renderRowMarkers();
        this.arrColWidth = this.setArrColumnWidth(); // use to calculate for resize column
        this.cellFormulaEl = $('cellStatus');
        
        this.addEvents();        

        
        this.setGridSize();
        this.createParticipantsMarkers();
        this.markActiveCell();
    },

    getMarkedValue: function(){
        return this.getActiveElement().get('text');    
    },
    
    isEditableFieldActive: function(){
        return this.editBox.getPosition().x > -1;
    },
    
    isLastCell: function(el){
        if(el) 
            return el.get('cell')==this.head.grid.rows[0].cells.length-1;
        else 
            return this.activeElement.cell==this.head.grid.rows[0].cells.length-1;
    },
    
    isFirstCell: function(el){
        if(el) 
            return el.get('cell')==1;
        else 
            return this.activeElement.cell==1;
    },
        
    isFirstRow: function(el){
        if(el) 
            return el.get('row')==1;
        else 
            return this.activeElement.row==1;
    },
    
    setActiveElement: function(row,cell,showMarker){
        if(row!=0) // don't mark header row
        {
            this.activeElement = {"row":row,"cell":cell};
            if(showMarker) this.markActiveCell();
            this.hideSelection();
        }
    },
    
    getActiveElement: function(){
        return this.getElement(this.activeElement['row'],this.activeElement['cell']);
    },

    getElement: function(row,cell){
        //check if the cell is in the frozen grid or the main grid
        if (row <= this.gridSize.frozenRows) 
            return this.head.grid.rows[row].cells[cell];    
        else
        {
            if(row > this.getCurrentMainRowCount())
            {
                while(row > this.getCurrentMainRowCount())
                    this.addRow();
            }
            return this.main.grid.rows[row-this.gridSize.frozenRows].cells[cell];
        }
    },
    
    getCurrentMainRowCount: function(bFrozen){ // count the existing rows
        if(bFrozen)
            return this.head.grid.rows;
        else
            return this.main.grid.rows.length - 2;
    },
    
    getElementByCaption: function(cellCaption){
        var cell = this.getCellFromAlphaNumericIndex(cellCaption);
        return this.getElement(cell.r,cell.c);
    },
    
    getActiveRow: function(row){
        //check if the cell is in the frozen grid or the main grid
        if (row <= this.gridSize.frozenRows) 
            return this.head.grid.rows[row];    
        else
            return this.main.grid.rows[row-this.gridSize.frozenRows];
    },

    markActiveCell: function(focusGrid){
        if(this.activeElement.row > this.gridSize.rows) // add row if needed
        {
            this.main.gridContainer.scrollTop = this.main.gridContainer.scrollHeight;
            this.addRow();
        }
        this.setRange(this.activeElement.row, this.activeElement.cell, this.activeElement.row, this.activeElement.cell); // save as range 
        targetCell = this.showCurrentActiveCell(); // this.grids.getElement('.divCurrentCell');
        targetCell.set('row',this.activeElement.row);
        targetCell.set('cell',this.activeElement.cell);
        var grid = this.getCurrentGrid(); 
        var position = this.getActiveElement().getPosition(grid)        
        targetCell.setPosition(position);
        var activeElementSize = this.getActiveElement().getSize();
        targetCell.setStyle('width',activeElementSize.x);
        targetCell.setStyle('height',activeElementSize.y);
        targetCell.setStyle('border-color',this.getColor('self'));
        //align scrolling
        var padding = 15;
        var xpadding = 100;
        if(!this.bRowInFreeze())
        {
            if(this.activeElement.row == (this.gridSize.frozenRows + 1))// if this is the first row in the header make sure the scroll is in the top
                this.main.gridContainer.scrollTop = 0;
            else if(position.y + activeElementSize.y + padding > parseInt(this.main.gridContainer.getStyle('height')))
                this.main.gridContainer.scrollTop = (position.y + activeElementSize.y + padding - 280)
        }
        if(position.x<=1)
            this.main.gridContainer.scrollLeft = 0;
        else if(position.x + activeElementSize.x + xpadding > window.getSize().x)
            this.main.gridContainer.scrollLeft = (position.x + activeElementSize.x + xpadding - window.getSize().x);
        
        // update the text box
        this.setCellFormula();
        if(focusGrid!=false)
            this.focusGrid();
        //this.positionScrolls(targetCell,position);
//                this.updateMarkerInState();
    },

    markOtherUserCell: function(participantId, row_, cell_){
        var row,cell,targetCell;
        participantId = $pick(participantId,this.viewerId);
        if(this.bLocalViewer(participantId)) // update state only for local participant change
        {        
            row = this.activeElement.row;
            cell = this.activeElement.cell;
            this.setRange(row, cell, row, cell); // save as range 
        }
        else
        {
            row=row_;
            cell=cell_;
        }
        targetCell = this.divHeaderContainer.getElement('div[forParticipant=' + participantId + ']');
        var bFade = false;
        if(targetCell == null)
        {
        	bFade = true;
            targetCell = this.addMarker(participantId,row,cell);
        }
        else if(targetCell.get('row')!=row || targetCell.get('cell')!=cell) // fade only if it's existing cell and it has changed
        	    bFade = true;
        targetCell.set('row',row);
        targetCell.set('cell',cell);
        targetCell.setPosition(this.getElement(row,cell).getPosition(this.head.grid));
        targetCell.setStyle('width',parseInt(this.getElement(row,cell).getSize().x));
        targetCell.setStyle('height',parseInt(this.getElement(row,cell).getSize().y));
        targetCell.setStyle('border-color',this.getColor(participantId));
//            if(this.bLocalViewer(participantId)) // update state only for local participant change
//                this.updateMarkerInState();
        
        if(!this.bLocalViewer(participantId) && bFade)
        {
        	var myEffect = new Fx.Morph(targetCell, {duration: 'long', transition: Fx.Transitions.Sine.quadIn});
    		myEffect.start({
    		    'opacity': [1, 0]
    		});
        }
    },

    bLocalViewer: function(participantId){
        return participantId==null || participantId==this.viewerId;
    },

    getColor: function(participantId){
        if(this.viewerId==participantId)
            return this.localParticipantColor;
        else
            return this.markerColor[participantId];
    },

    updateEditableInState: function(){
        if(wave.getState()!=null)
        {
            var editable = {};
            editable[this.viewerId] = {"row":this.activeElement.row,"cell":this.activeElement.cell};
            we.state.set('editables',editable,true);
        }   
    },

    removeEditableInState: function(participant){
        if(we.state!=null&&we.state.editables!=null)
            we.state.editables.unset(participant||this.viewerId,false);
    },

    updateMarkerInState: function(){
        if(wave.getState()!=null)
        {
            var marker = {};
            marker[this.viewerId] = {"row":this.activeElement.row,"cell":this.activeElement.cell};
            we.state.set('markers',marker,true);
        }   
    },
    
    getScrollLeft: function(){
        return $('divHeaderContainer').scrollLeft;
    },
    
    getScrollTop: function(){
        return $('divHeaderContainer').scrollTop;
    },
    
    showEditableField: function(keepValue,value){
        var position = this.getActiveElement().getPosition(this.editBoxContainer.parentNode);
        position.x += this.getScrollLeft();
        position.y += this.getScrollTop();
        this.editBoxContainer.setPosition(position);
        this.editBoxContainer.setStyle('width',parseInt(this.getActiveElement().getSize().x));
        this.editBoxContainer.setStyle('height',parseInt(this.getActiveElement().getSize().y) + 5);
        //this.editBox.value = this.getActiveElement().get('text');
        if(keepValue==true)
            this.editBox.value = this.getRealValue(this.getActiveElement());
        else
            this.editBox.value = value || '';
            
        this.editBox.set('editable',keepValue?'y':null);
        this.editBox.setAttribute('sourceFieldRow', this.activeElement.row);
        this.editBox.setAttribute('sourceFieldCell', this.activeElement.cell);
        setSelRange(this.editBox,this.editBox.value.length,this.editBox.value.length);

        this.updateEditableInState();
        this.adjustEditBox();

    },
    //recursive method to adjust the edit box to the right width in order not to have scroll bars
    adjustEditBox: function(){
        if(oSpreadsheet.editBox.getScroll().y > 0)
            oSpreadsheet.editBoxContainer.setStyle('width',parseInt(oSpreadsheet.editBoxContainer.getStyle('width')) + oSpreadsheet.editBox.getScroll().y + 5);
    },

    editBoxOnBlur: function(){
        if(this.isEditableFieldActive())
        {
            this.hideEditBox();
            this.removeEditableInState();
            this.setValue(this.sheet,
                parseInt(this.editBox.getAttribute('sourceFieldRow')),
                parseInt(this.editBox.getAttribute('sourceFieldCell')),
                this.editBox.value,null,null,null,false);
            this.selObj.formulaSelector.setStyle('display','none'); // hide formula selector if it's visible
            we.submitChanges();
        }
        this.markActiveCell();
    },
    
    deleteActiveElement: function(){
        this.setValue(this.sheet,this.activeElement.row,this.activeElement.cell,null,this.activeElement.row,this.activeElement.cell,this.currentRange,true);
        this.setFormulaField('');
    },
    
    setFormulaField: function(text){
        if(this.cellFormulaEl!=null)this.cellFormulaEl.value = text;
    },
    
    setValue: function(sheet,row,cell,value,toRow,toCell,range,submitChanges){
        if(range==null)
        {
            if(toRow==null)toRow=row;
            if(toCell==null)toCell=cell;
        }
        else
        {
            row = range.from.r;
            cell = range.from.c;
            toRow = range.to.r;
            toCell = range.to.c;
        }
        if(value=='')value = null; // in order to remove completely from the state
        var bFormula = (value!=null && value.substr(0,1) == '=');
        var formula;
        if(bFormula) // formula
        {
        	formula = value;
        	this.getElement(row,cell).set('formula',formula);
        	value = this.calculateValue(formula);
        }
        else
            this.getElement(row,cell).set('formula',null);
        
        for(var r=row;r<=toRow;r++)
        {
            for(var c=cell;c<=toCell;c++)
            {
            	we.state.set('data.' + r + '.' + c,{"v":value,"f":formula},false);
                this.getElement(r,c).set('text',value);
                this.onValueChange(r,c);
            }
        }
        if(submitChanges)
            we.submitChanges();
    },

    onValueChange: function(row,cell){
        this.updateDependencies(row,cell);
        this.alignRowHeaderHeight(row);
    },

    calculateValue: function(formula){
    	var uncoverFormula = formula.substr(1);
    	if(this.formulaHandler.isAdvancedFormula(formula))
    	{
    	    return this.formulaHandler.calculate(formula);
    	}
    	else
    	{
        	for(var currIndex=0;currIndex<formula.length;currIndex++)
        	{
        	    var cell = formula.substr(currIndex,1);
        		if(this.isAtoZ(cell)) // possible cell
        		{
        		    var row = this.getNumericPart(formula.substr(currIndex+1));
        		    if (row!="")
        		    {
        		        var totalCell = cell + row; // e.g. B24
        		        uncoverFormula = uncoverFormula.replace(totalCell,this.getElement(row,this.charToNumeric(cell)).get('text')); 
        		        currIndex = currIndex + totalCell.length; //advance the index to save calulation time
        		    }
        		}
        	}
        	try
        	{
        		return eval(uncoverFormula);
        	}
        	catch(e){return '#VALUE';}
        }
    },
    
    isNumeric: function(char){
        return String.charCodeAt(char)>=48 && String.charCodeAt(char)<=57;
    },

    isAtoZ: function(char){
        return String.charCodeAt(char)>63 && String.charCodeAt(char)<91;
    },
    //return the column D as 4
    charToNumeric: function(char){
        return String.charCodeAt(char)-64;
    },

    getNumericPart: function (str){
        var numericValue="";
        i=0;
        while(this.isNumeric(str.substr(i,1)) && i < str.length)
        {
            numericValue += str.substr(i,1);   
            i++; 
        }
        return numericValue;
    },
    
    setCss: function(sheet,styles,row,cell,toRow,toCell,range){
        if(range==null)
        {
            if(toRow==null)toRow=row;
            if(toCell==null)toCell=cell;
        }
        else
        {
            row = range.from.r;
            cell = range.from.c;
            toRow = range.to.r;
            toCell = range.to.c;
        }
        for(var r=row;r<=toRow;r++)
        {
            for(var c=cell;c<=toCell;c++)
            {
                we.state.set('data.' + r + '.' + c,{"s":styles},false);
                $H(styles).each(function(value, key) {
                    this.getElement(r,c).setStyle(key,value);    
                }.bind(this));
                
            }
        }
        we.submitChanges();
    },

    hideEditBox: function(){
        this.editBoxContainer.setPosition({x:-100,y:-100})
    },
    
    renderFrozenTable: function(){
        var tbody = this.head.grid.getElement('tbody');
        var oTrHeader = new Element('tr',{
            'class': 'hd'        
        }).injectInside(tbody);
        
        // set the first td
        var oFirstTd = new Element('td',{
                'styles':{
                    'border-bottom': '0px solid rgb(159, 194, 253)',
                    'width': '0pt'
                },
                'class':'hd first'
        }).injectInside(oTrHeader);
        
        new Element('input',{
                'styles':{
                    'width': '5px',
                    'height': '0px'
                },
                'value':'xx'
        }).injectInside(oFirstTd);
        
        // loop all header cells
        for (var i=0;i<this.gridSize.cells;i++)
        {
            var td = new Element('td',{
                    'styles':{
                        'border-bottom': '0px solid rgb(159, 194, 253)',
                        'width': '120px'
                    },
                    'class':'hd',
                    'text':String.fromCharCode(65 + i)
            }).injectInside(oTrHeader);
            
            var theResize = td.makeResizable({
                modifiers:{y:null}, // not to resize the height
                onComplete: function(el){
                    this.arrColWidth = this.setArrColumnWidth(); // use to calculate for resize column
                    //submit the change to the state
                    we.state.set('columns.' + el.cellIndex, {"w": el.getSize().x}, true);
                }.bind(this),
                onDrag: function(el){
                    $(this.main.grid.rows[0].cells[el.cellIndex]).setStyle('width',el.getStyle('width'));
                    this.markActiveCell();
                    this.alignRowHeaderHeight(this.activeElement.row);
                }.bind(this)
            });
        }
        
        //create data rows
        for (var r=0;r<this.gridSize.frozenRows;r++)
            this.addDataRow(tbody,false);
        
        //add the frozen bar
        //this.addFreezeBar();
        
    },
    
    setColumnWidth: function(cellIndex, width){
        $(this.head.grid.rows[0].cells[cellIndex]).setStyle('width',width);
        $(this.main.grid.rows[0].cells[cellIndex]).setStyle('width',width);
    },
    
    renderMainTable: function(){
        var tbody = this.main.grid.getElement('tbody');
        //position table 
        this.main.gridContainer.setStyle('top',this.head.gridContainer.getSize().y)
        //render the header row and the hidden rows
        this.addMainTableHeaderRow(tbody);
        
        //create data rows
        for (var r=0; r < this.gridSize.rows; r++)
            this.addDataRow(tbody, r < this.gridSize.frozenRows);
    },
    
    renderRowMarkers: function(){
        //position main marker table
        this.main.rowHeaderContainer.setStyle('top',this.head.gridContainer.getSize().y)
        
        for (var r=1; r <= (this.gridSize.rows); r++)
        {
            var bFrozenRow = this.bRowInFreeze(r);
            //var targetGrid = (this.bRowInFreeze(row)?this.frozRowHead:this.rowHead);
            if(bFrozenRow) 
                this.addRowMarker(this.frozRowHead,r,true);
            this.addRowMarker(this.rowHead,r,!bFrozenRow);
        }    
    },
    
    //add row to the main table
    addRow: function(){
        this.gridSize.rows++;
        tbody = this.main.grid.getElement('tbody');
        this.addRowMarker(this.rowHead, this.gridSize.rows, true);
        this.addDataRow(tbody,false);
        this.setMaximumSize();
        //gadgets.window.adjustHeight();
    },
    
    //add row to the main table
    addMainTableHeaderRow: function(tbody){
        var oTr = new Element('tr',{'class':'rShim'}).injectInside(tbody);
        //create the first cell for the height
        for (var c=0; c < this.gridSize.cells + 1; c++)
            new Element('td',{
                    'class': 'rShim',
                    'styles':{
                        'width': (c==0? '0pt': '120px')
                    }
                }).injectInside(oTr);
    },
    
    addFrozenRow: function(){
        this.gridSize.frozenRows++;
        tbody = this.head.grid.getElement('tbody');
        this.addRowMarker(this.frozRowHead, this.gridSize.frozenRows, true);
        this.addDataRow(tbody,true);
    },
    
    addRowMarker: function(table,index,visible){
            var tbody = table.getFirst();
            var tr = new Element('tr',{
                    'class': 'aupeezil',
                    'styles':{'display': (visible?'':'none')}
                }).injectInside(tbody);
                
            var td = new Element('td',{
                    'class': 'hd'
                }).injectInside(tr);
                
            var p = new Element('p',{
                    'styles': {
                        'overflow': 'hidden',
                        'line-height': '16px',
                        'height': '16px'
                    },
                    'text': index
                }).injectInside(td);
    },
    
    // add data row
    //<tr><td class="hd"><p style="height: 16px;">.</p></td><td/><td/><td/><td/><td/><td/><td/><td/><td/><td/><td/><td/><td/><td/></tr>
    addDataRow: function(tbody,bHidden){
        var oTr = new Element('tr').injectInside(tbody);
        if(bHidden)
            oTr.setStyle('display','none');
        //create the first cell for the height
        var oTd = new Element('td',{
            'class': 'hd'        
        }).injectInside(oTr);
    
        var oP = new Element('p',{
                'styles': {'height':'16px'}        
        }).injectInside(oTd);
    
        for (var c=0;c<this.gridSize.cells;c++)
            new Element('td').injectInside(oTr);
    },
    
    addFreezeBar: function(){
        var oTr = new Element('tr',{'id':'freezeBar'}).injectInside(this.head.grid.getElement('tbody'));
    
        for (var c = 0;c < this.gridSize.cells; c++)
            new Element('td',{
                       'class': 'freezeBar' + (c==0? ' rishon': ''),
                       'events': {
                            'mouseover': function(e){
                                  //this.showGhostFreezebar(); 
                            }.bind(this)
                       }
             }).injectInside(oTr);
    },
    
    addMarker: function(participant,row,cell)
    {
        //create marker divs
        return new Element('div',{   
            'forParticipant':participant.replace(/\./,'_'),
            //'participantName':participant.getDisplayName(),
            'class':'divCurrentCell',
            'row':row,
            'cell':cell,
            'styles':{
                   'width': '122px', 'height':'19px','top': '37px', 'left': '-1px', 'z-index':'95'
            }}).injectInside(this.divHeaderContainer);
    },

    updateMarkers: function(markers)
    {
    	if(markers)
    	{
		//get all the markers
		markers.filterKeys($not($begins('$'))).each(function(value,key){
			if(key!="_cursorPath")
			    this.markActiveCell(key, value.row, value.cell);
		    }.bind(this));
	    }
    },

    updateEditables: function(editables)
    {
        // remove all existing editing cells
        $$('div.divOtherUser').each(
            function(a){
                a.destroy()
            });
        
        if(editables)
        {
		    //remove all previous editable
            $$('.divOtherCursorEditing').each(function(item){
                item.removeClass('divOtherCursorEditing');
            });
            //get all the markers
            editables.filterKeys($not($begins('$'))).each(function(value,participant){
               if(value != null && participant != '_cursorPath')
               {
                if(participant!=this.viewerId && value.cell!=null && value.row != null)
                {
                	 //get marker 
                	var target = this.bRowInFreeze(value.row)?this.head:this.main;
                	var targetCell = this.getOtherUserEditingDiv(participant);//$$('div.divOtherUser')[0];
               	    target.gridContainer.adopt(targetCell);
                	
                	targetCell.addClass('divOtherUserEditing');
                	targetCell.setStyle('display','');
                    targetCell.set('row',value.row);
                    targetCell.set('cell',value.cell);
                    targetCell.setPosition(this.getElement(value.row,value.cell).getPosition(target.gridContainer));
                    targetCell.setStyle('width',parseInt(this.getElement(value.row,value.cell).getSize().x));
                    targetCell.setStyle('height',parseInt(this.getElement(value.row,value.cell).getSize().y));
            	    this.releaseOtherUserField.delay(60000,this,participant); // if the user is in editable state more than x amount of time remove it from the state
            	//if (marker!=null)
            	//    marker.addClass('divOtherCursorEditing');
                }
            }
            }.bind(this));
        }
    },

    updateColumns: function(columns)
    {
    	if(columns)
    	{
    		//get all the markers
    		columns.filterKeys($not($begins('$'))).each(function(value,key){
    			if(key!="_cursorPath")
    			    this.setColumnWidth(key, value.w);
		    }.bind(this));
		    //make sure the active cell is refreshed and the row headers aligned
            this.markActiveCell(false);
            this.alignRowHeaderHeight(this.activeElement.row);
            this.arrColWidth = this.setArrColumnWidth(); // update the column width
 
	}

    },


    //<div class="divOtherUser" style="border-color: rgb(16, 150, 24); border-width: 2px; top: 55px; left: 362px; width: 122px; height: 19px;display:none;"></div>
    getOtherUserEditingDiv: function(participant)
    {
        return new Element('div',{
            'class': 'divOtherUser',
            'styles': {
                    'border-color': 'rgb(16, 150, 24)',
                    'border-width': '2px' 
            },
            'forParticipant': participant
        })
    },

    updateActions: function(actions)
    {
    	if(actions)
    	{
		// iterate through rows
		    actions.filterKeys($not($begins('$'))).each(function(row,rowkey){
    			if(rowkey!="_cursorPath")
    			{
    				// iterate through cells
    				row.filterKeys($not($begins('$'))).each(function(cell,cellkey){     
    					if(cellkey!="_cursorPath")
    					{
    						cell.filterKeys($not($begins('$'))).each(function(value,typekey){     
    								switch(typekey)
    								{
    									case 'v':
    									      this.getElement(rowkey,cellkey).set('text',value);
    									      //if there is a function to this cell attach it
    									      if(cell.f!=null)
    									        this.getElement(rowkey,cellkey).set('formula',cell.f);
                                              
    									break;	
    									case 's':
    									      this.getElement(rowkey,cellkey).set('styles',value);
    									break;	
    								}
    						}.bind(this))
    					}                 			
    				}.bind(this))
    		    }
		    }.bind(this));
		    
		    this.alignAllRowHeader();
	    }
    },

    getMarker: function(participant)
    {
        return this.divHeaderContainer.getElement('div[forParticipant=' + participant + ']');
    },

    createParticipantsMarkers: function()
    {
        //for each participant make sure there is marker
        if(this.participants!=null)
        {
            this.participants.each(function(key,value){
                    if(key.getId()!=this.viewerId)
                    {
                        var oElement = this.divHeaderContainer.getElement('.forParticipant=' + value);
                        if (oElement==null)
                            this.addMarker(key);
                    }
                }.bind(this));
        }
    },

    btnBold_onclick: function()
    {
         var element = this.getElement(this.activeElement.row,this.activeElement.cell);
         var css = element.getStyle('font-weight');
         if(css.toLowerCase()=='bold')
              this.setCss(this.sheet,{'font-weight':'normal'},this.activeElement.row,this.activeElement.cell, null, null, this.currentRange);
         else
              this.setCss(this.sheet,{'font-weight':'bold'},this.activeElement.row,this.activeElement.cell, null, null, this.currentRange);
    },
    
    btnItalic_onclick: function()
    {
         var element = this.getElement(this.activeElement.row,this.activeElement.cell);
         var css = element.getStyle('font-style');
         if(css.toLowerCase()=='italic')
              this.setCss(this.sheet,{'font-style':'normal'},this.activeElement.row,this.activeElement.cell, null, null, this.currentRange);
         else
              this.setCss(this.sheet,{'font-style':'italic'},this.activeElement.row,this.activeElement.cell, null, null, this.currentRange);
    },
    
    positionElement: function(targetCell,e){
            var keyPressed = e.key;
            var cell = targetCell.get('cell');
            var row = targetCell.get('row');
            switch(keyPressed)
            {
                case 'left':
                    if(!this.isFirstCell(targetCell))
                        cell--;
                    break;
                 case 'right':
                    if(this.isLastCell(targetCell))
                    {
                        cell=1;
                        row++;
                    }
                    else
                        cell++;

                    break;
                 case 'up':
                    if(!this.isFirstRow(targetCell))
                        row--;
                    break;
                 case 'down':
                        row++;
                    break;
            }
            targetCell.set('row',row);
            targetCell.set('cell',cell);
            targetCell.setPosition(this.getElement(row,cell).getPosition(this.matrix));
            targetCell.setStyle('width',parseInt(this.getElement(row,cell).getSize().x));
            targetCell.setStyle('height',parseInt(this.getElement(row,cell).getSize().y));
    },
    
    showColorPalette: function(e,type){

        if($('colorTable') == null)
        {
            var table = new Element('table',{
                    id:'colorTable',
                    styles:{'width':'170px','height':'170px','position':'absolute','top':e.client.y,'left':e.client.x,'z-index':'300'},
                    cellspacing:0,
                    cellpadding:0,
                    border:0,
                    'type':type
                }).injectInside(document.body);
            
            var tbody = new Element('tbody').injectInside(table);
            
            for(var r=0;r<this.colorPalatte.length;r++)
            {
                var row = new Element('tr').injectInside(tbody);
                for(var c=0;c<this.colorPalatte[0].length;c++)
                {
                     //create cell
                     var cell = new Element('td',{
                            'styles':{
                                'background':'#' + this.colorPalatte[r][c],
                                'border': '1px solid white'
                            },
                            'color': this.colorPalatte[r][c]
                     }).injectInside(row);
                     cell.addEvents({
                                'mouseout': function(){
                                    this.setStyle('border', '1px solid white');
                                }.bind(cell),
                                'mouseover': function(){
                                    this.setStyle('border', '1px inset black');
                                }.bind(cell),
                                'click': function(color){
                                    this.selectColor('#' + color);
                                }.bind(this,this.colorPalatte[r][c])
                            });
                     var img = new Element('img',{
                            'styles': {
                                'width': '12',
                                'height': '12'
                            },
                            'src': 'http://dwido.thewe.net/spreadsheet/spacer.gif'
                        }).injectInside(cell);
                    
                }
            }
            
        } 
        else 
        {
            $('colorTable').set('type',type);
            $('colorTable').setStyle('display','');
        }
    },
    
    selectColor: function(color){
            $('colorTable').setStyle('display','none');
            if($('colorTable').get('type')=='highlightColor')
                this.setCss(this.sheet,{'background':color},this.activeElement.row,this.activeElement.cell,null,null,this.currentRange);
            else if($('colorTable').get('type')=='textColor')
                this.setCss(this.sheet,{'color':color},this.activeElement.row,this.activeElement.cell,null,null,this.currentRange);

            //this.getActiveElement().setStyle('background',color);
    },
    
    getRealValue: function(el){
        var formula = el.get('formula');
        if(formula!=null && formula!= '')
            return formula;
        else
            return el.get('text');
    },
    
    selectRange: function(rowOffset, cellOffset, bShift){
        var position,toRow,toCell;
        var bInit = (this.selObj.selection.getStyle('display')=='none');
        if(bInit)
            this.selObj.selection.setStyle('display','');
        
        toRow = this.currentRange.to.r + rowOffset;
        toCell = this.currentRange.to.c + cellOffset;
        
        if(!bShift)
        {
            fromRow = toRow;
            fromCell = toCell;
        }
        else
        {
            fromRow = this.currentRange.from.r;
            fromCell = this.currentRange.from.c;
        }

        //this.currentRange = {'from':{'r':Math.min(toRow,fromRow), 'c': Math.min(toCell,fromCell)},'to':{'r': Math.max(toRow,fromRow), 'c':Math.max(toCell,fromCell)}};
        this.setRange(fromRow, fromCell, toRow, toCell);
        this.selObj.selection.setPosition(this.getElement(this.currentRange.from.minr, this.currentRange.from.minc).getPosition(this.matrix)); // position selection div

        position = this.getElement(this.currentRange.to.maxr + 1, this.currentRange.to.maxc + 1).getPosition(this.matrix); // get the next cell,row location which marks the end of the selection
        //it's already showing and positioned just view it
        this.selObj.selection.setStyle('width', Math.abs(position.x - parseInt(this.selObj.selection.style.left)));
        this.selObj.selection.setStyle('height', Math.abs(position.y - parseInt(this.selObj.selection.style.top)));
//        this.selObj.selection.set('toRow', toRow);
//        this.selObj.selection.set('toCell', toCell);
    },
    
/*    initSelection: function(){
        this.selObj.selection.setStyle('display','');
        this.selObj.selection.setPosition(this.getActiveElement().getPosition(this.matrix)); // position selection div
        this.selObj.selection.set('toCell', this.activeElement.cell);
        this.selObj.selection.set('toRow', this.activeElement.row);
        position = this.getElement(this.activeElement.row + 1,this.activeElement.cell + 1).getPosition(this.matrix); // get the next cell,row location which marks the end of the selection
        this.selObj.selection.setStyle('width', Math.abs(position.x - parseInt(this.selObj.selection.style.left)));
        this.selObj.selection.setStyle('height', Math.abs(position.y - parseInt(this.selObj.selection.style.top)));
        this.selObj.selection.set('toRow', this.activeElement.row);
        this.selObj.selection.set('toCell', this.activeElement.cell);
    },
  */  
    hideSelection: function(){
        if(this.selObj.selection.style.display!='none')
        {
            //this.currentRange = null;
            this.selObj.selection.setStyle('display','none');
            this.selObj.selection.set('toRow',null);
            this.selObj.selection.set('toCell',null);
        }
    },
    
    removeAllHighlights: function (){
    	this.frozRowHead.getElements('.on').removeClass('on');
    	this.rowHead.getElements('.on').removeClass('on');
    	this.head.grid.rows[0].getElements('.on').removeClass('on');
    },
    
    setHighlights: function (fromRow, fromCell, toRow, toCell){
    	this.removeAllHighlights();
    	for(var r = Math.min(fromRow, toRow); r <= Math.max(fromRow, toRow); r++)
    		this.getContainerRowHead(r).rows[r].addClass('on');

    	for(var c = Math.min(fromCell,toCell); c <= Math.max(fromCell,toCell); c++)
    		this.head.grid.rows[0].cells[c].addClass('on');
    },
    
    getContainerRowHead: function(row){
        row = $pick(row,this.activeElement.row);
        return (this.bRowInFreeze(row))?this.frozRowHead:this.rowHead;
    },
    
    getContainerGrid: function(row){
        row = $pick(row,this.activeElement.row);
        return (row<=this.gridSize.frozenRow)?this.head.grid:this.main.grid;
    },
    
    bRowInFreeze: function(row){
        row = $pick(row || this.activeElement.row);
        return (row <= this.gridSize.frozenRows);
    },
    
    setRange: function(fromRow, fromCell, toRow, toCell){
        this.currentRange = {
            'from':{'r': fromRow, 'c': fromCell, 'minr': Math.min(toRow, fromRow), 'minc': Math.min(toCell, fromCell)},
            'to':{'r': toRow, 'c': toCell, 'maxr': Math.max(toRow, fromRow), 'maxc': Math.max(toCell, fromCell)}
        };
        
    	this.setHighlights(fromRow, fromCell, toRow, toCell);
    	this.updateCellLocation();
    },

    getRangeCaption: function(){
        if(this.bMultiRange())
            return this.getRangeAlphaNumericIndex();
        else
            return this.getCellAlphaNumericIndex(this.currentRange.from.r,this.currentRange.from.c);
    },
        
    bMultiRange : function(){
        return ((this.currentRange.from.r != this.currentRange.to.r) || (this.currentRange.from.c != this.currentRange.to.c));
    },

    alignRowHeaderHeight:    function(row){
        //if(this.getContainerGrid(row).rows[row].getSize().y>parseInt(this.getContainerRowHead(row).rows[row].getFirst().getStyle('height')))debugger;
        this.getContainerRowHead(row).rows[row].getFirst().setStyle('height',this.getActiveRow(row).getSize().y);
        this.adjustGadgetHeight();
    },
    
    adjustGadgetHeight: function(){
        if($('heightAutoAdjust').checked)
            gadgets.window.adjustHeight();
    },

    alignAllRowHeader:    function(){
        for (var row = 1; row < this.gridSize.rows; row++)
            this.alignRowHeaderHeight(row);
    },
    
    // when a cell value is changed it is checking if any other cell is depended on that to update the formula
    updateDependencies: function(row,cell){
        var cellAN = this.getCellAlphaNumericIndex(row,cell);
        //check for all the cells that contain this cell in the formula     
        this.grids.getElements('[formula]').each(function(a,b){
                if(a.get('formula').indexOf(cellAN)!=-1) // the changed cell is part of the formula so update the formula
                    this.refreshFormula(a);
            }.bind(this));
    },
    
    refreshFormula: function(el){
       	value = this.calculateValue(el.get('formula'));
        el.set('text',value);
    	we.state.set('data.' + el.parentNode.rowIndex + '.' + el.cellIndex,{"v": value},true);
    },
    
    // e.g row-5 cell-3 return E3
    getCellAlphaNumericIndex: function(row,cell){
        return String.fromCharCode(64+parseInt(cell)).toUpperCase() + row;
    },

    // column 3 --> 'C'
    getColumnAlphaNumeric: function(column){
        return String.fromCharCode(64+parseInt(column)).toUpperCase();
    },
    
    // e.g 'E3' --> {r: 5, c: 3}
    getCellFromAlphaNumericIndex: function(cell){
        //find the location where the numbers starts
        var i = 0;
        var retV = {};
        while(this.isAtoZ(cell.substr(i)) && i < cell.length)
        {i++};
        
        if(i > 0)
            return {'r': parseInt(cell.substr(i)), 'c': this.charToNumeric(cell.substr(0,i))}
        
    },

    // returns E3:F8
    getRangeAlphaNumericIndex: function(){
        return this.getCellAlphaNumericIndex(this.currentRange.from.minr,this.currentRange.from.minc) + ":" + this.getCellAlphaNumericIndex(this.currentRange.to.maxr,this.currentRange.to.maxc);
    },
    
    // E7:F8 -> {'from':{'r': 5, 'c': 7},'to':{'r': 6, 'c': 8}}; 
    getRangeFromAlphaNumericIndex: function(txtRange){
        var arrCells = txtRange.split(':'); 
        var fromCell = this.getCellFromAlphaNumericIndex(arrCells[0]);
        var toCell = this.getCellFromAlphaNumericIndex(arrCells[1]);
        return {'from':{'r': fromCell.r, 'c': fromCell.c, 'minr': Math.min(fromCell.r,toCell.r), 'minc': Math.min(fromCell.c,toCell.c)}, 'to':{'r': toCell.r, 'c': toCell.c, 'maxr': Math.max(fromCell.r,toCell.r), 'maxc': Math.max(fromCell.c,toCell.c)}};
    },
    
    // to update the Location of the active cell in the left cell on the bar on top of the grid
    updateCellLocation: function(range){ 
        /*if(this.currentRange.from.c==this.currentRange.to.c&&this.currentRange.from.r==this.currentRange.to.r)
            this.cellLocationTd.set('text',this.getCellAlphaNumericIndex(this.currentRange.from.r,this.currentRange.from.c))*/
    },
    
    setGridSize: function(){
        var windowSize = window.getSize();
        /*debugger;
        var yHeader = parseInt($('divHeaderContainer').getStyle('height'));
        var totalHeight = windowSize.y-x;
        var divTopheight = totalHeight-16;
        var mainHeight = totalHeight - yHeader - 3;
        //$('divHeaderContainer').setStyle('width',windowSize.x-57);
        this.grids.setStyle('height',totalHeight);
        this.matrix.getElement('.divTop').setStyle('height',divTopheight);
        this.matrix.getElement('.divMainArea').setStyle('height',mainHeight);
        this.rowHead.setStyle('height',mainHeight);*/
        /* width */
        this.main.gridContainer.setStyle('width',windowSize.x - 40);
        this.head.gridContainer.setStyle('width',windowSize.x - 41);
    },
    
    showCurrentActiveCell: function() {
        if(this.bRowInFreeze())
        {
            this.main.currentCell.setStyle('display', 'none');
            return this.head.currentCell.setStyle('display', '');
        }
        else
        {
            this.head.currentCell.setStyle('display', 'none');
            return this.main.currentCell.setStyle('display', '');
        }
    },
    
    getCurrentGridContainer: function(){
        return this.bRowInFreeze()?this.head.gridContainer:this.main.gridContainer;    
    },

    getCurrentGrid: function(){
        return this.bRowInFreeze()?this.head.grid:this.main.grid;    
    },
    
    showFormulaDialog: function(){
		new MochaUI.Window({
		    title: 'Insert a function',
			id: 'formulaDialog',
			loadMethod: 'json',
			content: $('tblInsertFunction').clone(true),
			width: 350,
			height: 170,
			x: 85,
			y: 50
		});
		
		var rows = $('formulaDialog').getElements('.funcTable tr');
		rows.each(function(row){
		    row.addEvent('mouseover',function(a,row){
                $('formulaDialog').getElement('.funcSyntax').set('text',row.get('funcsyntax'));
                //this.cellFormulaEl.value = '=' + row.get('funcname') + '()';
                this.getActiveElement().set('text', '=' + row.get('funcname') + '()');
		    }.bindWithEvent(this,row));
		    row.addEvent('click',function(e,row){
		        MochaUI.closeWindow($('formulaDialog'));
		        //this.cellFormulaEl.setCaretPosition(this.cellFormulaEl.value.indexOf('()') + 1);
		        this.showEditableField(false, '=' + row.get('funcname') + '(');
		        //this.editBox.setCaretPosition(this.editBox.value.indexOf('(') + 1);
		        //setSelRange(this.cellFormulaEl, this.cellFormulaEl.value.indexOf('()') + 1, this.cellFormulaEl.value.indexOf('()') + 1)
		    }.bindWithEvent(this,row));
		}.bind(this))
    },
    
    setCellFormula: function(text){
        //this.cellFormulaEl.value = (text?text:$pick(this.getActiveElement().get('formula'),this.getActiveElement().get('text')));
        this.cellFormulaEl.set('text',(text?text:$pick(this.getActiveElement().get('formula'),this.getActiveElement().get('text'))));
    },
    
    isRange: function(text){
        return text.contains(':');    
    },
    
    // "A7:B8" --> ['A7','A8','B7','B8']
    parseRange: function(range){
        var arr = [];
        var range = this.getRangeFromAlphaNumericIndex(range);
        for (var r = range.from.minr; r <= range.to.maxr; r++)
        {
            for (var c = range.from.minc; c <= range.to.maxc; c++)
            {
                arr[arr.length] = this.getColumnAlphaNumeric(c) + r;
            }
        }
        return arr;
    },
    
    focusGrid: function(){
        $('xxFocus').focus();    
    },
    
    //calculate the total height of elements within the container to adjust height
    setMaximumSize: function(){
        //totalHeight = document.body.getElement('ul.markdown').getSize().y + $('headerMarker').getSize().y + this.head.gridContainer.getSize().y + this.matrix.getElement('.divMainArea').getSize().y;
        if($('heightAutoAdjust').checked)
        {        
            totalHeight = this.head.gridContainer.getSize().y + this.matrix.getElement('.divMainArea').getSize().y;
            this.grids.setStyle('height',totalHeight);
            if($('heightAutoAdjust').checked)
                gadgets.window.adjustHeight();
        }
    },
    
    setArrColumnWidth: function(){
        var arrWidth = [];
        for(var col=0; col < this.head.grid.rows[0].cells.length; col++)
            arrWidth[col] = this.head.grid.rows[0].cells[col].getSize().x;
        return arrWidth;
    },
    
    getCursorResizeInfo: function(x)
    {
        var accumulatedWidth = 0;
        var columnIndex = 0;
        var bLoop = true;
        var ret;
        while(columnIndex < this.arrColWidth.length && bLoop)
        {
            var nextAcumulation = accumulatedWidth + this.arrColWidth[columnIndex];
            if(x < nextAcumulation)
            {
                ret = {'c':columnIndex,'x' : accumulatedWidth, 'offsetleft': x - accumulatedWidth,'offsetright': nextAcumulation - x, 'width': this.arrColWidth[columnIndex]}
                bLoop = false;
            }
            else
                accumulatedWidth = nextAcumulation 
            
            columnIndex++;
        }
        return ret;
    },
    
    positionColumnResizer: function(){
        var resizer = this.grids.getElement('.columnResizer');
        resizer.setPosition(this.head.grid.getPosition(this.grids));
        return resizer;
    },
    
    addEvents: function(){
        this.head.grid.addEvent('click',function(e){
            this.setActiveElement(e.target.parentNode.rowIndex,e.target.cellIndex,true);
        }.bind(this));

        this.main.grid.addEvent('click',function(e){
            this.setActiveElement(e.target.parentNode.rowIndex,e.target.cellIndex,true);
        }.bind(this));

        this.editBox.addEvent('blur',function(){
            this.editBoxOnBlur();
        }.bind(this));

        this.editBox.addEvent('keydown',function(e){
                if(this.editBox.value.substr(0,1)=="=") // formula
                {
                    if(e.key=='right' || e.key=='left' || e.key=='up' || e.key=='down')
                    {
                        /*if(this.selObj.formulaSelector.getStyle('display')=='none')
                        {
                            this.selObj.formulaSelector.set('cell',this.activeElement.cell);
                            this.selObj.formulaSelector.set('row',this.activeElement.row);
                            this.selObj.formulaSelector.setStyle('display','');
                            this.editBox.origValue = this.editBox.value; // remember the original value
                        }
                        this.positionElement(this.selObj.formulaSelector,e);*/
                        if(this.editBox.get('rangeSelection')!='active')
                        {
                            this.editBox.origValue = this.editBox.value;
                            this.editBox.set('rangeSelection','active');
                        }
                        
                        switch(e.key)
                        {
                            case 'right':
                                this.selectRange(0, 1, e.shift);
                                break;
                            case 'left':
                                this.selectRange(0, -1, e.shift);
                                break;
                            case 'up':
                                this.selectRange(-1, 0, e.shift);
                                break;
                            case 'down':
                                this.selectRange(1, 0, e.shift);
                                break;
                        }
                        
                        //this.editBox.value = this.editBox.origValue + (String.fromCharCode(64+parseInt(this.selObj.formulaSelector.get('cell'))).toUpperCase()) + this.selObj.formulaSelector.get('row');
                        this.editBox.value = this.editBox.origValue + this.getRangeCaption();
                        return false;
                    }
                    else
                    {
                        if(e.code!=16) // for shift do nothing
                        {
                            this.editBox.set('rangeSelection',null);
                            this.selObj.formulaSelector.setStyle('display','none');
                        }
                    }
                    this.adjustEditBox();
                }
                if (e.key=='enter')
                {
                    this.editBox.fireEvent('blur');
                    this.focusGrid();
                    //this.markActiveCell();
                }
                setTimeout(this.adjustEditBox,0);
                
        }.bind(this));
        
        this.editBox.addEvent('keyup',function(e){this.setCellFormula(this.editBox.value)}.bind(this));

        this.editBoxContainer.addEvent('keydown',function(e){ // to keep the cursor in the edit box when user clicks arrows
                if(e.key=='right'||e.key=='left'||e.key=='up'||e.key=='down')
                    if (this.editBox.get('editable')) // stop propagation so the arrows will navigate within the cell
                        e.stopPropagation();
        }.bind(this));

        
        /*this.cellFormulaEl.addEvent('keydown',function(e){
                if(e.key=='right'||e.key=='left'||e.key=='up'||e.key=='down')
                {
                    
                    // mark selection
                    switch(e.key)
                    {
                        case 'right':
                            this.selectRange(0, 1, e.shift);
                            break;
                        case 'left':
                            this.selectRange(0, -1, e.shift);
                            break;
                        case 'up':
                            this.selectRange(-1, 0, e.shift);
                            break;
                        case 'down':
                            this.selectRange(1, 0, e.shift);
                            break;
                    }
                    
                    this.cellFormulaEl.insertAtCursor(this.getRangeCaption());
                    return false;
                }
                else if(e.key=='enter')
                {
                    if(this.formulaHandler.isAdvancedFormula(this.cellFormulaEl.value))
                    {
                        this.getActiveElement().set('text', this.cellFormulaEl.value);
                        this.getActiveElement().set('formula', this.cellFormulaEl.value);
                        this.refreshFormula(this.getActiveElement());
                        this.focusGrid();
                    }
                    else
                        this.getActiveElement().set('text',this.cellFormulaEl.value);

                    return false;
                }
                else this.selObj.formulaSelector.setStyle('display','none');
        }.bind(this));*/
        
        
        // keys events
        document.addEvent('keydown', function(e) {
            var stopPropagation = false;
            if(e.shift && (e.key=='right' || e.key=='left' || e.key=='up' || e.key=='down'))
            {
/*                if(this.selObj.selection.getStyle('display')=='none') // it's hidden initialize it
                    this.initSelection();*/

                // mark selection
                switch(e.key)
                {
                    case 'right':
                        this.selectRange(0, 1, e.shift);
                        break;
                    case 'left':
                        this.selectRange(0, -1, e.shift);
                        break;
                    case 'up':
                        this.selectRange(-1, 0, e.shift);
                        break;
                    case 'down':
                        this.selectRange(1, 0, e.shift);
                        break;
                }
                return false;
            } 
            else if (e.key=='right' || (e.key=='tab' && !e.shift))
            {
                //check if this the last item on the row
                if(this.isLastCell())
                {
                    this.activeElement.cell=1;
                    this.activeElement.row++;
                }
                else
                    this.activeElement.cell++;

                this.editBox.fireEvent('blur');
                this.markActiveCell();
                this.hideSelection(); //hide selection div if visible
                stopPropagation = true;
            }
            else  if (e.key=='left' || (e.key=='tab' && e.shift))
            {
                //check if this the last item on the row
                if(!this.isFirstCell())
                    this.activeElement.cell--;
                    this.editBox.fireEvent('blur');
                //this.markActiveCell();
                this.hideSelection(); //hide selection div if visible
                stopPropagation = true;
            }
            else  if (e.key=='up')
            {
                //check if this the last item on the row
                if(this.isFirstRow())
                    return false;
                else
                {
                    this.activeElement.row--;
                    //this.markActiveCell();
                    this.editBox.fireEvent('blur');
                }
                this.hideSelection(); //hide selection div if visible
                stopPropagation = true;
            }
            else  if (e.code==36) // home
            {
                this.activeElement.cell = 1;
                this.hideEditBox(); // hide the edit box and clear it
                this.markActiveCell();
                this.hideSelection(); //hide selection div if visible
                stopPropagation = true;
            }
            else  if (e.key=='down')
            {
                //check if this the last item on the row
                this.activeElement.row++;
                this.editBox.fireEvent('blur');
                this.hideSelection(); //hide selection div if visible
                stopPropagation = true;
            }
            else  if (e.key=='delete')
            {
                this.deleteActiveElement();
                this.markActiveCell();
                stopPropagation = true;
            }
            else  if (e.key=='f2'||e.key=='enter')
            {
                if(e.target.tagName.toLowerCase()!='textarea')
                {
                    this.showEditableField(true);
                    //this.markActiveCell();
                    stopPropagation = true;
                }
                else 
                    this.editBox.fireEvent('blur');
                    
            }
            else if (e.code == 18) //enter
            {
                this.markActiveCell();
                return true;
            }
            else if (e.code == 16) // shift
            {
                //do nothing
            }
            else if (e.key == 'esc')
            {
                this.hideEditBox(); // hide the edit box and clear it
                this.editBox.value = '';
            }
            else
            {
                if(!this.isEditableFieldActive())
                    this.showEditableField();
            }
            
            if(e.target.tagName.toLowerCase()!='textarea')
            {
                if(stopPropagation)
                    return false
            }
            else
                return true;
                
        }.bind(this));

        //text area where functions are edited
        /*this.cellFormulaEl.addEvent('keydown',function(e){
                e.stopPropagation();
            })
        */
        this.matrix.getElement('.divMainArea').addEvent('scroll',function(){
            this.head.gridContainer.scrollLeft = this.main.gridContainer.scrollLeft;
            this.main.rowHeaderContainer.scrollTop = this.main.gridContainer.scrollTop;
        }.bind(this))

        //color palette
        $('btnTextColor').addEvent('click',function(e){
            this.showColorPalette(e,'textColor');
        }.bind(this));

        //color palette
        $('btnHighlight').addEvent('click',function(e){
            this.showColorPalette(e,'highlightColor');
        }.bind(this));

        window.addEvent('resize',function(){
            this.setGridSize();
        }.bind(this));

        // Column resize events
        this.head.grid.rows[0].addEvent('mousemove',function(e){
            var headerOffsetLeft = this.head.grid.getPosition(this.grids).x;
            var resizeInfo = this.getCursorResizeInfo(e.client.x - this.head.grid.getPosition(this.grids).x);
            
            if(resizeInfo.offsetright < this.resizeMargine)
                this.head.grid.rows[0].setStyle('cursor','col-resize');
            else 
                this.head.grid.rows[0].setStyle('cursor','default');
                
        }.bind(this));
        
        var myDrag = new Drag(this.columnResizer, {
            modifiers: {x: 'left',y: null},
            onStart: function(el){
                //                ret = {'c':columnIndex,'x' : accumulatedWidth, 'offsetleft': x - accumulatedWidth,'offsetright': nextAcumulation - x, 'width': this.arrColWidth[columnIndex]}
                //this.columnResizer.columnEl = this.head.grid.rows[0].cells[this.columnResizer.resizeInfow.c];
            }.bind(this),
            onDrag: function(el){
                //this.columnResizer.columnEl.setStyle('width',)
            },
            onComplete: function(el){
            }
        });
   
        // add drug functionality to freeze bar
        new Drag(this.ghostFreezebar, {
            modifiers: {x: null},
            onStart: function(el){

            }.bind(this),
            onDrag: function(el){

            },
            onComplete: function(el){
            }
        });

    },

    showGhostFreezebar: function(){
        var freezebar = this.getFreezeBar();
        this.ghostFreezebar.setPosition(freezebar.getPosition(this.grids));
        this.ghostFreezebar.setStyle('width',freezebar.getSize().x);
        this.ghostFreezebar.setStyle('height',freezebar.getSize().y);
    },
    
    getFreezeBar: function(){
        return this.head.grid.rows[this.head.grid.rows.length - 1];
    },
    
    releaseOtherUserField: function(participant){
        var el = this.grids.getElement(".divOtherUserEditing[forParticipant='" + participant + "']");
        //remove the user from the state
        this.removeEditableInState(participant);
        we.submitChanges();
    }
    
/*    // "A7:B8" --> [{r:7,c:1,name:'A7'},{r:8,c:1,name:'A8'},{r:7,c:2,name:'B7'},{r:8,c:2,name:'B8'}]
    parseRange: function(range){
        var arr = [];
        var range = this.getRangeFromAlphaNumericIndex(range);
        for (var r = range.from.minr; r <= range.to.maxr; r++)
        {
            for (var c = range.from.minc; c <= range.to.maxc; c++)
            {
                arr[arr.length] = {'r': r, 'c': c, 'name': this.getColumnAlphaNumeric(c) + r};
            }
        }
        return arr;
    }    
*/    
});

    function setSelRange(inputEl, selStart, selEnd) { 
     if (inputEl.setSelectionRange) { 
      inputEl.focus(); 
      inputEl.setSelectionRange(selStart, selEnd); 
     } else if (inputEl.createTextRange) { 
      var range = inputEl.createTextRange(); 
      range.collapse(true); 
      range.moveEnd('character', selEnd); 
      range.moveStart('character', selStart); 
      range.select(); 
     } 
    }

    function initialize() {
        oSpreadsheet.userId = String.fromCharCode(97 + Math.round(Math.random() * 25)) + String.fromCharCode(97 + Math.round(Math.random() * 25)) + String.fromCharCode(97 + Math.round(Math.random() * 25));
    }

    function initialState() {
        return {value: 0}
    }
    
    function stateUpdated(state) {
          //alert(wave.getViewer());
        if(oSpreadsheet && oSpreadsheet.viewerId==null && wave.getViewer()!=null) //set the div for the active viewer
        {  
            oSpreadsheet.viewerId = wave.getViewer().getId().replace(/\./g, "_");
            oSpreadsheet.head.currentCell.setAttribute('forParticipant',oSpreadsheet.viewerId);
            oSpreadsheet.main.currentCell.setAttribute('forParticipant',oSpreadsheet.viewerId);
        }
        
        if(oSpreadsheet && oSpreadsheet.grids!=null) //the page was loaded
        {
           oSpreadsheet.updateMarkers(state.markers);
           oSpreadsheet.updateEditables(state.editables);
           oSpreadsheet.updateActions(state.data);
           oSpreadsheet.updateColumns(state.columns);
        }
    }
    
    function participantStateUpdated(state) {
        oSpreadsheet.participants = state;
    }
    //gadgets.util.registerOnLoadHandler(init);
    var oSpreadsheet;
    
    window.addEvent('domready', function(){oSpreadsheet = new Spreadsheet()});
