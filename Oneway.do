//Last update 13/05/20

//Example
/*
Premier argument : variables
Deuxième argument : path et filename
Trosième argument : lower threshold des variables exclues pour le calcul des pourcentages
Quatrième argument : upper threshold des variables exclues pour le calcul des pourcentages
DOneway "GDP Fraise var2" "C:/Machin/Truc.xlsx" 90 99
*/

//NB: Stata gets progressively slower as you accumulate putexcel commands. 

capture program drop Oneway //Cool thing so you don't have to erase your program each time you wanna update it
program Oneway
version 14.2
args VarRowList Topic1 FileName LowThres UpperThres varname Missing
	local CellCol = 1
	local CellRow = 3 //Deux premiers row: noms des vars colonnes, etc.
	local CellFreq = char(65+`CellCol') + string(`CellRow'+1)
	local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1)
	/*
	capture confirm file "$gtable/Tris à plat/nul" // check if `name' subdir exists
	if _rc { // _rc will be >0 if it doesn't exist
		!md "$gtable/Tris à plat/"
	}*/
	quietly putexcel set "`FileName'", sheet("Output") replace //Set excel directory 
	
	quietly putexcel A1:A2, merge
	quietly putexcel B1 = "No.", vcenter right font(Calibri, 10)
	quietly putexcel B2 = "(N = `=_N')", vcenter right font(Calibri, 10)
	quietly putexcel C2 = "%", vcenter right font(Calibri, 10)
	quietly putexcel A1:C1, border(top)
	quietly putexcel A2:C2, border(bottom)
	
	if "`Topic1'" != "" {
	quietly putexcel A3 = "`Topic1'", vcenter txtwrap bold font(Calibri, 10)
	local CellRow = 4 //Deux premiers row: noms des vars colonnes, etc.
	local CellFreq = char(65+`CellCol') + string(`CellRow'+1)
	local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1)
	}
	
	foreach VarRow in `VarRowList' {
	
		/*//MOD1
		Old stuff I used for another study. But it's too aggressive anyway
		forvalues i = 90/99 {
			quietly replace `VarRow' = `i' if `VarRow' == -`i'
		} //It's just to get the NR/NSP and stuff at the end, if it's negative
		*/
		
		local RowValueLabel : value label `VarRow' //Value label storing the labels (it's actually the name of the variable here)
		levelsof `VarRow', local(RowLevels) //Stores the numeric categories of VarRow to the local RowLevels
		
			//Partie pourcentages 
		tab `VarRow' if !inrange(`VarRow',`LowThres',`UpperThres') , matcell(frequencies) matrow(row) //Tabulate command, mat`stuff' creates useful matrices		
		mata: st_matrix("colperc", (st_matrix("frequencies")  :/ colsum(st_matrix("frequencies")))*100) //Matrice pourcentages colonne
		matrix Col_perc = colperc[1...,.]
		
		if "`Missing'" != "m" {
			//Partie fréquences
			tab `VarRow', matcell(frequencies) matrow(row) //Tabulate command, mat`stuff' creates useful matrices
			scalar NbRow = r(r) //Nombre lignes
			matrix Col_freq = frequencies[1...,.]
		}
		
		if "`Missing'" == "m" {
			//Partie fréquences
			quietly tab `VarRow', matcell(frequencies) matrow(row) //Tabulate command, mat`stuff' creates useful matrices
			scalar NbRowLabels = r(r) //Nombre lignes labels
			tab `VarRow', m matcell(frequencies) matrow(row) //Tabulate command, mat`stuff' creates useful matrices
			scalar NbRow = r(r) //Nombre lignes total
			matrix Col_freq = frequencies[1...,.]
		}
		
		//In Excel
		
		if "`varname'" != "varname" {
			local VarRowLabel : variable label `VarRow'
			quietly putexcel A`CellRow' = "`VarRowLabel'", vcenter txtindent(1) txtwrap nformat(number_sep) font(calibri, 10)  //Label Question 	
		}
		
		if "`varname'" == "varname" {
		quietly putexcel A`CellRow' = "`VarRow'", vcenter txtindent(1) txtwrap nformat(number_sep) font(calibri, 10)  //Label Question 	
		}
		
		quietly putexcel `CellFreq' = matrix(Col_freq), vcenter right font(calibri, 10)  //Colonne fréquences
		quietly putexcel `CellPerc' = matrix(Col_perc), vcenter right nformat("0.00") font(calibri, 10)  //Colonne pourcentages
		
		if "`Missing'" != "m" {
			forvalues MatRow = 1/`=NbRow' { //Matrice des labels
				local CellLabelRow = string(`CellRow'+`MatRow')
				local RowValueLabelNum = word("`RowLevels'", `MatRow')
				local CellContents : label `RowValueLabel' `RowValueLabelNum'
				quietly putexcel A`CellLabelRow' = "`CellContents'", /*italic*/ vcenter txtindent(2) txtwrap font(calibri, 10) 
			}
		}
		
		if "`Missing'" == "m" {
			forvalues MatRow = 1/`=NbRowLabels' { //Matrice des labels
				local CellLabelRow = string(`CellRow'+`MatRow')
				local RowValueLabelNum = word("`RowLevels'", `MatRow')
				local CellContents : label `RowValueLabel' `RowValueLabelNum'
				quietly putexcel A`CellLabelRow' = "`CellContents'", /*italic*/ vcenter txtindent(2) font(calibri, 10) 
			}
			
			if `=NbRow' > `=NbRowLabels' {
				local MissingValueTextRow = `CellRow' + `=NbRow'
				quietly putexcel A`MissingValueTextRow' = "Valeurs manquantes", vcenter txtindent(2) font(calibri, 10) 
			}
		}
		
		local CellRow = `CellRow' + r(r) + 1 //Update CellRow for new Variable
		local CellFreq = char(65+`CellCol') + string(`CellRow'+1)
		local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1)
		
		/*//MOD1
		forvalues i = 96/99 {
			quietly replace `VarRow' = `i' if `VarRow' == -`i'
		}
		*/
	}
	local CellRow = `CellRow' - 1
	quietly putexcel A`CellRow':C`CellRow', border(bottom)
end
