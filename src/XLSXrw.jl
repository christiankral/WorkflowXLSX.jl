__precompile__(true)
module XLSXrw

    using XLSX
    export
        rowcol2cell,
        XLSXsheetname,
        XLSXreadvar,
        XLSXwritevar


    # Convert "missing" to ""
    function missing2empty(data)
        # Convert elements "missing" by ""
        for col in eachcol(data)
           replace!(col,missing => "")
        end
        return data
    end

    # Convert row and cell to Excel cell string
    function rowcol2cell(row,col)
        if col > 702
            error("rowcol2cell: Cannot convert rows > 702")
        end
        col1 = mod(col - 1, 26) + 1
        col2 = fld(col - 1, 26) + 1
        if col2 == 1
            cell = ""
        else
            cell = Char(col2 + 63)
        end
        cell = cell * Char(col1 + 64) * string(row)
        return cell
    end

    # Rename name of first sheet
    function XLSXsheetname(file, sheetname)
        XLSX.openxlsx(file, mode="rw") do xf
            sheet = xf[1]
            XLSX.rename!(sheet, sheetname)
        end
    end

    function XLSXreadvar(file, var)
        xi = XLSX.readxlsx(file)
        # Sheets
        sheets = XLSX.sheetnames(xi)
        # First sheet = input sheet
        sh = xi[sheets[1]]
        # Read entire sheet and replace "missing" by ""
        data = missing2empty(shi[:])
        # Indexes of valid entries
        ind = findall(x->x == var, data)
        if length(ind) == 0
            error("Variable \'" * var *  "\' is not present in any XLSX cell")
        elseif length(ind) > 1
            error("Variable \'" * var *  "\' is present more than once in XSLX cell")
        else
            if size(data,2) < ind[1][2]+1
                error("Numeric value of variable \'" * var * "\' is not present in XSLSX file")
            end
            v = data[ind[1][1],ind[1][2]+1]
            if !(typeof(v) == Int64 || typeof(v) == Float64)
                error("Numeric value of variable \'" * var * "\' is not present in XSLSX file")
            end
            if size(data,2) < ind[1][2]+2
                println("WARNING: Unit of variable \'" * var * "\' ist not present in XSLX file")
                return v
            else
                u = data[ind[1][1],ind[1][2]+2]
                if u == "" || u == 1
                    return v
                else
                    return v * uparse(u)
                end
            end
        end

    end

    function XLSXwritevar(file, var, qnt, unt; dig=6)
        if unt == 1
            # If qnt == 1 then substitute conversion to
            # 1 = m/m = Meter/Meter
            printuln(var, qnt, m/m)
        else
            # Use propagated unit unt
            printuln(var, qnt, unt)
        end
        XLSX.openxlsx(file, mode="rw") do xf
            sheet = xf[1]
            # Load entire content of sheet and
            # replace elements "missing" by ""
            data = missing2empty(sheet[:])
            # Indexes of each valid entry
            ind = findall(x->x == var, data)
            if length(ind) == 0
                error("Variable \'" * var *  "\' is not present in any XLSX cell")
            elseif length(ind) > 1
                error("Variable \'" * var *  "\' is present in more than once XLSX cell")
            end
            # Convert Cartesian index to row and column
            row = ind[1][1]
            col = ind[1][2]
            # Convert quantity 'qnt' to destionation unit 'unt'
            if unt == 1
                QNT = uconvert(m/m, qnt)
            else
                QNT = uconvert(unt, qnt)
            end
            VAL = round(QNT / (1 * unt), digits = dig)
            # Cell of variable
            cell_var = rowcol2cell(row,col)
            # Cell of value
            cell_val = rowcol2cell(row,col+1)
            # Cell of unit
            cell_unt = rowcol2cell(row,col+2)
            # Write value and unit
            sheet[cell_val] = VAL
            if unt != 1
                sheet[cell_unt] = string(unt)
            end
        end
    end
end
