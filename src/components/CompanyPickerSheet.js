// Company picker opened from Profile when the user has 2+ allowed companies.
// Tapping a row hands it to onPick; the caller asks for confirmation before
// switching. Adapted from Tools_rental_management's BranchPickerSheet.
import React from 'react';
import { Modal, View, Text, StyleSheet, TouchableOpacity, ScrollView } from 'react-native';
import { MaterialIcons } from '@expo/vector-icons';
import { COLORS, FONT_FAMILY } from '@constants/theme';

const NAVY = COLORS.primaryThemeColor;

const CompanyPickerSheet = ({ visible, companies = [], currentId = null, onPick, onClose }) => (
  <Modal
    visible={!!visible}
    transparent
    animationType="fade"
    statusBarTranslucent
    onRequestClose={onClose}
  >
    <View style={styles.backdrop}>
      <View style={styles.sheet}>
        <View style={styles.headerRow}>
          <View style={{ flex: 1 }}>
            <Text style={styles.title}>Choose company</Text>
            <Text style={styles.subtitle}>The app shows registers, products and customers of this company.</Text>
          </View>
          <TouchableOpacity
            style={styles.closeBtn}
            onPress={onClose}
            hitSlop={{ top: 12, bottom: 12, left: 12, right: 12 }}
            activeOpacity={0.7}
          >
            <MaterialIcons name="close" size={22} color="#666" />
          </TouchableOpacity>
        </View>

        <ScrollView style={styles.list} contentContainerStyle={{ paddingBottom: 8 }} showsVerticalScrollIndicator={false}>
          {companies.map((c) => {
            const isActive = c.id === currentId;
            return (
              <TouchableOpacity
                key={c.id}
                style={[styles.row, isActive && styles.rowActive]}
                activeOpacity={0.85}
                onPress={() => onPick(c)}
              >
                <View style={[styles.radio, isActive && styles.radioActive]}>
                  {isActive ? <View style={styles.radioDot} /> : null}
                </View>
                <Text style={[styles.rowName, isActive && styles.rowNameActive]} numberOfLines={1}>
                  {c.name}
                </Text>
                {isActive ? <MaterialIcons name="check-circle" size={20} color="#fff" /> : null}
              </TouchableOpacity>
            );
          })}
        </ScrollView>
      </View>
    </View>
  </Modal>
);

const styles = StyleSheet.create({
  backdrop: { flex: 1, backgroundColor: 'rgba(0,0,0,0.5)', justifyContent: 'center', paddingHorizontal: 24 },
  sheet: {
    backgroundColor: '#fff', borderRadius: 20, paddingTop: 22, paddingHorizontal: 18, paddingBottom: 14,
    shadowColor: '#000', shadowOpacity: 0.15, shadowRadius: 14, shadowOffset: { width: 0, height: 6 }, elevation: 8,
  },
  headerRow: { flexDirection: 'row', alignItems: 'flex-start', marginBottom: 14 },
  closeBtn: {
    width: 32, height: 32, borderRadius: 16, alignItems: 'center', justifyContent: 'center',
    backgroundColor: '#F0F0F2', marginLeft: 8,
  },
  title: { fontSize: 18, fontFamily: FONT_FAMILY.urbanistBold, color: NAVY, marginBottom: 2 },
  subtitle: { fontSize: 12, fontFamily: FONT_FAMILY.urbanistMedium, color: '#888' },
  list: { maxHeight: 380 },
  row: {
    flexDirection: 'row', alignItems: 'center', paddingVertical: 13, paddingHorizontal: 14, borderRadius: 14,
    marginBottom: 8, backgroundColor: '#F8F9FA', borderWidth: 1, borderColor: '#E8E8E8',
  },
  rowActive: { backgroundColor: NAVY, borderColor: NAVY },
  radio: {
    width: 22, height: 22, borderRadius: 11, borderWidth: 2, borderColor: NAVY,
    alignItems: 'center', justifyContent: 'center',
  },
  radioActive: { borderColor: '#fff' },
  radioDot: { width: 10, height: 10, borderRadius: 5, backgroundColor: '#fff' },
  rowName: { flex: 1, marginLeft: 12, fontSize: 15, fontFamily: FONT_FAMILY.urbanistSemiBold, color: '#333' },
  rowNameActive: { color: '#fff' },
});

export default CompanyPickerSheet;
