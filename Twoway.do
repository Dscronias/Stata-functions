capture  program drop Twoway
program Twoway
version 14.2
args VarColList VarRowList FileName LowThreshold UpThreshold varname row
foreach VarCol in `VarColList' {


	//Set File
	quietly putexcel set "`FileName' `VarCol'.xlsx", sheet("`VarCol'") replace //Set excel directory 

////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	
	
	//Headers
	quietly: tab `VarCol', matcell(ColVarFreq)
	matrix ColVarFreq = matrix(ColVarFreq)'
		//Obs for each category
	forvalues col = 1/`=r(r)' {
		matrix ColVarFreq_`col' = ColVarFreq[1..., `col']
		scalar ColVarFreq_`col' = round(ColVarFreq_`col'[1,1])
	}
		//Total obs for the variable
	scalar Total = r(N)
		//Category label
	local ColValueLabel : value label `VarCol' //Value label storing the labels (it's actually the name of the variable here)
	levelsof `VarCol', local(ColLevels) //Stores the numeric categories of VarCol to the local RowLevels
		//Put category labels and number of observations

	local CurrentColLeft = char(66)
	local CurrentColRight = char(66+1)
	quietly: tab `VarCol'
	forvalues Col = 1/`=r(r)' { //Matrice des labels
			local ColValueLabelNum = word("`ColLevels'", `Col')
			local CellContents : label `ColValueLabel' `ColValueLabelNum'
			//Category label
			quietly putexcel `CurrentColLeft'2:`CurrentColRight'2, merge
			quietly putexcel `CurrentColLeft'2 = "`CellContents'", right vcenter bold font(Calibri, 10)
			
			//Category Obs
			quietly putexcel `CurrentColLeft'3:`CurrentColRight'3, merge
			quietly putexcel `CurrentColLeft'3 = "(N = `=ColVarFreq_`Col'')", right vcenter font(Calibri, 10)
			
			//N, %
			quietly putexcel `CurrentColLeft'4 = "N", vcenter right font(Calibri, 10)
			if "`row'" != "row" {
				quietly putexcel `CurrentColRight'4 = "Col. %", vcenter right font(Calibri, 10)
			}
			if "`row'" == "row" {
				quietly putexcel `CurrentColRight'4 = "Row %", vcenter right font(Calibri, 10)
			}
			
			//Next category
			local CurrentColLeft = char(66+`Col'*2)
			local CurrentColRight = char(66+`Col'*2+1)		
	}
			//Total
	scalar NTotal = round(e(N_pop))
	quietly putexcel `CurrentColLeft'2:`CurrentColRight'2, merge
	quietly putexcel `CurrentColLeft'2 = "Total", right vcenter bold font(Calibri, 10)
	quietly putexcel `CurrentColLeft'3:`CurrentColRight'3, merge
	quietly putexcel `CurrentColLeft'3 = "(N = `=Total')", right vcenter font(Calibri, 10)
	quietly putexcel `CurrentColLeft'4 = "N", right vcenter font(Calibri, 10)
	quietly putexcel `CurrentColRight'4 = "%", right vcenter font(Calibri, 10)
			
	local VarColLabel : variable label `VarCol'
	quietly putexcel B1:`CurrentColRight'1, merge  //Label Question 	
	if "`varname'" == "varname" {
		quietly putexcel B1 = "`VarCol'", right bold vcenter txtwrap font(Calibri, 10)  //Label Question 	
	}
	if "`varname'" != "varname" {
		quietly putexcel B1 = "`VarColLabel'", right bold vcenter txtwrap font(Calibri, 10)  //Label Question 	
	}
	quietly putexcel A1:`CurrentColRight'1, border(top)
	quietly putexcel A4:`CurrentColRight'4, border(top)
	
	
////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	
	//Main
	
	local CellCol = 1
	local CellRow = 5 //Deux premiers row: noms des vars colonnes, etc.
	local CellFreq = char(65+`CellCol') + string(`CellRow'+1)
	local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1)
	
	
	foreach VarRow in `VarRowList' {
		local RowValueLabel : value label `VarRow' //Value label storing the labels (it's actually the name of the variable here)
		levelsof `VarRow', local(RowLevels) //Stores the numeric categories of VarRow to the local RowLevels
		
			//Partie pourcentages 
		if "`row'" != "row" {	
			tab `VarRow' `VarCol' if !inrange(`VarRow',`LowThreshold',`UpThreshold'), column matcell(frequencies) matrow(row) matcol(col) chi //Tabulate command, mat`stuff' creates useful matrices		
			mata: st_matrix("colperc", (st_matrix("frequencies")  :/ colsum(st_matrix("frequencies")))*100) //Matrice pourcentages colonne
		}
		if "`row'" == "row" {	
			tab `VarRow' `VarCol' if !inrange(`VarRow',`LowThreshold',`UpThreshold'), row matcell(frequencies) matrow(row) matcol(col) chi //Tabulate command, mat`stuff' creates useful matrices		
			mata: st_matrix("colperc", (st_matrix("frequencies")  :/ rowsum(st_matrix("frequencies")))*100) //Matrice pourcentages ligne
		}
		scalar NbCol = r(c) //Nombre colonnes
		scalar NbCol_2 = r(c)*2 //Nombre colonnes *2
		scalar Chi = r(p) //Chi square
		//local ChiCell = char(65+`=NbCol_2'+3) + string(`CellRow'+1)
		forvalues col = 1/`=NbCol' {
			matrix Col`col'_perc = colperc[1..., `col']
		} //Chaque colonne dans une matrice différente
		
		//Partie fréquences
		if "`row'" != "row" {
			tab `VarRow' `VarCol', column matcell(frequencies) matrow(row) matcol(col) //Tabulate command, mat`stuff' creates useful matrices
		}
		if "`row'" == "row" {
			tab `VarRow' `VarCol', row matcell(frequencies) matrow(row) matcol(col) //Tabulate command, mat`stuff' creates useful matrices
		}
		scalar NbRow = r(r) //Nombre lignes
		forvalues col = 1/`=NbCol' {
			matrix Col`col'_freq = frequencies[1..., `col']
		} //Chaque colonne dans une matrice différente
		
		//Total pourcentages
		local TotPercCell = char(65+`=NbCol_2'+2) + string(`CellRow'+1)
		quietly tab `VarRow' if !inrange(`VarRow',`LowThreshold',`UpThreshold') & `VarCol' != ., matcell(frequencies)
		mata: st_matrix("ColTot_Perc", (st_matrix("frequencies")  :/ colsum(st_matrix("frequencies")))*100) //Matrice pourcentages colonne
		
		//Total fréquences
		local TotFreqCell = char(65+`=NbCol_2'+1) + string(`CellRow'+1)
		quietly tab `VarRow' if `VarCol' != ., matcell(ColTot_Freq)
		
		//In Excel
		if "`varname'" != "varname" {
			local VarRowLabel : variable label `VarRow'
			quietly putexcel A`CellRow' = "`VarRowLabel'", vcenter bold txtwrap nformat(number_sep) font(Calibri, 10)  //Label Question 	
			quietly putexcel A`CellRow':`CurrentColRight'`CellRow', merge
		}
			//Input row variable name
		if "`varname'" == "varname" {
		quietly putexcel A`CellRow' = "`VarRow'", vcenter bold txtwrap nformat(number_sep) font(Calibri, 10)  //Label Question
		quietly putexcel A`CellRow':`CurrentColRight'`CellRow', merge
		}
		quietly putexcel `TotFreqCell' = matrix(ColTot_Freq), vcenter right nformat("0") font(Calibri, 10) 
		quietly putexcel `TotPercCell' = matrix(ColTot_Perc), vcenter right nformat("0.00") font(Calibri, 10) 	 	
		//quietly putexcel `ChiCell' = `=Chi', vcenter hcenter nformat("0.0000") font(calibri, 10) fpattern(solid, white, white) //Test Chi-sq
		forvalues MatCol = 1/`=NbCol' { //Matrice des données
			quietly putexcel `CellFreq' = matrix(Col`MatCol'_freq), vcenter right nformat("0") font(Calibri, 10) //Colonne fréquences
			quietly putexcel `CellPerc' = matrix(Col`MatCol'_perc), vcenter right nformat("0.00") font(Calibri, 10) //Colonne pourcentages
			local CellCol = `CellCol' + 2 //Paire de colonnes suivante 
			local CellFreq = char(65+`CellCol') + string(`CellRow'+1) //Update
			local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1) //Update
		}
		forvalues MatRow = 1/`=NbRow' { //Matrice des labels
			local CellLabelRow = string(`CellRow'+`MatRow')
			local RowValueLabelNum = word("`RowLevels'", `MatRow')
			local CellContents : label `RowValueLabel' `RowValueLabelNum'
			quietly putexcel A`CellLabelRow' = "`CellContents'", vcenter txtindent(1) txtwrap font(Calibri, 10) 
		}
		
		local CellRow = `CellRow' + r(r) + 1 //P-value
		quietly putexcel A`CellRow' = "P-value:", vcenter txtindent(1) font(Calibri, 10)
		if `=Chi' > 0.05 {	
			quietly putexcel B`CellRow' = `=Chi', right vcenter nformat("0.00") font(Calibri, 10)
		}
		else if `=Chi' < 0.001 {
			quietly putexcel B`CellRow' = "< 0.001", right vcenter bold font(Calibri, 10)
		}
		else {
			quietly putexcel B`CellRow' = `=Chi', right vcenter nformat("0.000") bold font(Calibri, 10)
		}
		local CellRow = `CellRow' + 1 //New Variable
		local CellCol = 1 //This needs to reset for each new variable
		local CellFreq = char(65+`CellCol') + string(`CellRow'+1) //Update
		local CellPerc = char(65+`CellCol'+1) + string(`CellRow'+1) //Update
	}
	//Footer
	local CellRow = `CellRow' - 1
	quietly putexcel A`CellRow':`CurrentColRight'`CellRow', border(bottom)
}

end
