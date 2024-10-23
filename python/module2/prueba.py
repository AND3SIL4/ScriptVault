def update_exception_file(self, 
                            data: Dict,
                            inconsistencies: pd.DataFrame) -> None:
        """Actualiza el archivo de excepciones con los nuevos consecutivos"""
        try:
            # Cargar el archivo existente
            book = load_workbook(self.exception_file)
            sheet = book["CONSECUTIVO SAP"]
            
            # Actualizar valores
            sheet.cell(row=1, column=2).value = data['initial_consecutivo']  # Consecutivo inicial
            sheet.cell(row=1, column=3).value = data['final_consecutivo']    # Consecutivo final
            
            # Actualizar lista de consecutivos pendientes
            row = 1
            for consecutivo in data['nuevos_consecutivos']:
                if consecutivo in inconsistencies['CONSECUTIVO_FROM_CONSECUTIVO'].values:
                    row += 1
                    sheet.cell(row=row, column=1).value = consecutivo
            
            # Guardar cambios
            book.save(self.exception_file)
            return True
        except Exception as e:
            print(f"Error al actualizar el archivo: {str(e)}")
            return Falsedef update_exception_file(self, 
                            data: Dict,
                            inconsistencies: pd.DataFrame) -> None:
        """Actualiza el archivo de excepciones con los nuevos consecutivos"""
        try:
            # Cargar el archivo existente
            book = load_workbook(self.exception_file)
            sheet = book["CONSECUTIVO SAP"]
            
            # Actualizar valores
            sheet.cell(row=1, column=2).value = data['initial_consecutivo']  # Consecutivo inicial
            sheet.cell(row=1, column=3).value = data['final_consecutivo']    # Consecutivo final
            
            # Actualizar lista de consecutivos pendientes
            row = 1
            for consecutivo in data['nuevos_consecutivos']:
                if consecutivo in inconsistencies['CONSECUTIVO_FROM_CONSECUTIVO'].values:
                    row += 1
                    sheet.cell(row=row, column=1).value = consecutivo
            
            # Guardar cambios
            book.save(self.exception_file)
            return True
        except Exception as e:
            print(f"Error al actualizar el archivo: {str(e)}")
            return False