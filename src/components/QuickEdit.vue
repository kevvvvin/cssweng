<template>
  <q-select
    bg-color="white"
    filled
    v-model="currPlayer"
    :options="playerNames"
    label="Quick Edit"
  />

  <q-table
    v-if="currPlayer"
    :rows="players[currPlayer.value].bets"
    :title="titleFn"
    :rows-per-page-options="[]"
    row-key="day"
    dense
    :pagination="{ rowsPerPage: 20 }"
    >
    <template v-slot:body="props">
      <q-tr :props="props">
        <q-td key="day" :props="props">
          {{ props.row.day }}
          <q-popup-edit v-model="props.row.day" auto-save v-slot="scope">
            <q-input
              v-model="scope.value"
              dense
              autofocus
              counter
              @keyup.enter="scope.set"
            />
          </q-popup-edit>
        </q-td>
        <q-td key="amount" :props="props">
          {{ props.row.amount }}
          <q-popup-edit v-model="props.row.amount" auto-save v-slot="scope">
            <q-input
              v-model="scope.value"
              dense
              autofocus
              counter
              @keyup.enter="scope.set"
            />
          </q-popup-edit>
        </q-td>
        <q-td key="team" :props="props">
          {{ props.row.team }}
          <q-popup-edit v-model="props.row.team" auto-save v-slot="scope">
            <q-input
              v-model="scope.value"
              dense
              autofocus
              counter
              @keyup.enter="scope.set"
            />
          </q-popup-edit>
        </q-td>
        <q-td key="result" :props="props">
          {{ props.row.result }}
          <q-popup-edit v-model="props.row.result" auto-save v-slot="scope">
            <q-input
              v-model="scope.value"
              dense
              autofocus
              counter
              @keyup.enter="scope.set"
            />
          </q-popup-edit>
        </q-td>
      </q-tr>
    </template>
  </q-table>
</template>
<style scoped>
q-table {
  width: 400px;
}
</style>
<script>
import { defineComponent, ref } from "vue";
import { useQuasar } from "quasar";

export default defineComponent({
  name: "QuickEdit",
  props: {},
  setup() {
    const $q = useQuasar();
    return {
      titleFn({ currPlayer }) {
        return `${currPlayer.label} 's bets [t: ${players[currPlayer.value].tong}% | c: ${players[currPlayer.value].comm}%]`;
      },
    };
  },
  async mounted() {
    this.darkMode = $q.dark.isActive;
  },
});
</script>